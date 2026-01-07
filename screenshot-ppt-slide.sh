#!/bin/bash
#set -x  # Enable debug logging
exec 2> /tmp/ppt_trace.log  # Log all commands to trace file

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title Screenshot PPT Slide
# @raycast.mode silent

# Optional parameters:
# @raycast.icon 📸
# @raycast.packageName PowerPoint Tools
# @raycast.description Screenshot current PowerPoint slide and name it with course code + slide title

# Documentation:
# @raycast.author Jordon Ockey

# Get PowerPoint window info and slide title
temp_image="/tmp/ppt_slide_capture_$$.png"
applescript_file="/tmp/ppt_script_$$.scpt"

# Create AppleScript file
cat > "$applescript_file" <<'APPLESCRIPT'
tell application "System Events"
    tell process "Microsoft PowerPoint"
        if not (exists window 1) then
            return "ERROR: No PowerPoint window found"
        end if
        
        set frontmost to true
        delay 0.2
        
        set windowTitle to name of window 1
        set windowPos to position of window 1
        set windowSize to size of window 1
        set winX to item 1 of windowPos
        set winY to item 2 of windowPos
        set winW to item 1 of windowSize
        set winH to item 2 of windowSize
        
        -- Calculate slide area (less aggressive for edge detection)
        -- We'll capture a larger area and let Swift detect the actual slide boundaries
        -- More aggressive top crop since toolbar is always there
        set slideX to winX + 200
        set slideY to winY + 150
        set slideW to winW - 220
        set slideH to winH - 200
    end tell
end tell

-- Try to get slide title and number from PowerPoint
set slideTitle to ""
set slideNum to ""
try
    tell application "Microsoft PowerPoint"
        set activePres to active presentation
        set slideNum to slide index of slide of view of active window as string
        
        -- Try to get title from slide
        try
            set currentSlide to slide (slideNum as integer) of activePres
            set shapeCount to count of shapes of currentSlide
            if shapeCount > 0 then
                repeat with i from 1 to shapeCount
                    try
                        set shp to shape i of currentSlide
                        set shapeText to content of text range of text frame of shp
                        if shapeText is not "" and shapeText is not missing value then
                            set slideTitle to shapeText
                            exit repeat
                        end if
                    on error
                    end try
                end repeat
            end if
        on error
        end try
    end tell
on error
end try

return windowTitle & "|" & slideX & "|" & slideY & "|" & slideW & "|" & slideH & "|" & slideTitle & "|" & slideNum
APPLESCRIPT

ppt_info=$(osascript "$applescript_file")
rm -f "$applescript_file"

# Check for errors
if [[ "$ppt_info" == *"ERROR"* ]]; then
    echo "$ppt_info"
    exit 1
fi

# Parse the info
window_title=$(echo "$ppt_info" | cut -d'|' -f1)
slide_x=$(echo "$ppt_info" | cut -d'|' -f2)
slide_y=$(echo "$ppt_info" | cut -d'|' -f3)
slide_w=$(echo "$ppt_info" | cut -d'|' -f4)
slide_h=$(echo "$ppt_info" | cut -d'|' -f5)
ppt_title=$(echo "$ppt_info" | cut -d'|' -f6)
slide_num=$(echo "$ppt_info" | cut -d'|' -f7)

# Extract course code (Smart Logic)
course_code=$(python3 -c "
import sys, re

name = sys.argv[1]
# Remove extensions
name = re.sub(r'\.pptx?$', '', name, flags=re.IGNORECASE)

# Remove MED- prefix (case insensitive)
name = re.sub(r'^med[-_ ]?', '', name, flags=re.IGNORECASE)

if not name:
    print('PPT')
    sys.exit()

# Take first 2 chars
res = name[:2]

# Look at 3rd and 4th chars
for i in range(2, len(name)):
    if i >= 4: break
    c = name[i]
    
    # Stop at separators
    if c in ' -_.,': break
    
    # Check case consistency with previous char
    prev = name[i-1]
    
    # If 2nd char was upper, continue only if current is upper
    if prev.isupper() and c.isupper():
        res += c
    # If 2nd char was lower, continue only if current is lower
    elif prev.islower() and c.islower():
        res += c
    else:
        break

print(res)
" "$window_title")

# Capture a larger area for edge detection
screencapture -R "${slide_x},${slide_y},${slide_w},${slide_h}" -x "$temp_image"

if [ $? -ne 0 ] || [ ! -f "$temp_image" ]; then
    echo "ERROR: Screenshot failed"
    exit 1
fi

# Auto-detect slide boundaries and crop tightly  
crop_script="/tmp/crop_slide_$$.swift"
cat > "$crop_script" <<'SWIFT'
#!/usr/bin/env swift
import Foundation
import Cocoa

guard CommandLine.arguments.count >= 3 else {
    exit(1)
}

let imagePath = CommandLine.arguments[1]
let outputPath = CommandLine.arguments[2]

guard let image = NSImage(contentsOfFile: imagePath) else {
    exit(1)
}

guard let cgImage = image.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
    exit(1)
}

let width = cgImage.width
let height = cgImage.height

// Convert image to pixel data
guard let colorSpace = CGColorSpace(name: CGColorSpace.sRGB),
      let context = CGContext(data: nil, width: width, height: height, bitsPerComponent: 8, bytesPerRow: width * 4, space: colorSpace, bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else {
    exit(1)
}

context.draw(cgImage, in: CGRect(x: 0, y: 0, width: width, height: height))
guard let pixelData = context.data?.assumingMemoryBound(to: UInt8.self) else {
    exit(1)
}

import Darwin

func logError(_ message: String) {
    fputs("\(message)\n", stderr)
}

// PPT background gray detector
func isBackgroundGray(_ pixel: UnsafePointer<UInt8>, _ offset: Int) -> Bool {
    let r = Int(pixel[offset])
    let g = Int(pixel[offset + 1])
    let b = Int(pixel[offset + 2])
    let isNeutral = abs(r - g) < 8 && abs(g - b) < 8 && abs(r - b) < 8
    let brightness = (r + g + b) / 3
    return isNeutral && brightness >= 25 && brightness <= 70
}

// Generic dark pixel helper (used when counting dark rows)
func isDarkPixel(_ pixel: UnsafePointer<UInt8>, _ offset: Int) -> Bool {
    return isBackgroundGray(pixel, offset)
}

// Non-gray slide pixel (any color that's not PPT chrome)
func isSlidePixel(_ pixel: UnsafePointer<UInt8>, _ offset: Int) -> Bool {
    if isBackgroundGray(pixel, offset) {
        return false
    }
    let r = Int(pixel[offset])
    let g = Int(pixel[offset + 1])
    let b = Int(pixel[offset + 2])
    let brightness = (r + g + b) / 3
    let maxDiff = max(abs(r - g), abs(g - b), abs(r - b))
    return brightness > 75 || maxDiff > 20
}

// Detect top edge by finding first non-gray dominant row
var topEdge = 0
for y in 0..<height {
    let startOffset = (y * width) * 4
    let sampleStart = width / 6
    let sampleEnd = width - sampleStart
    var slidePixels = 0
    var grayPixels = 0
    for x in sampleStart..<sampleEnd {
        let offset = startOffset + (x * 4)
        if isSlidePixel(pixelData, offset) {
            slidePixels += 1
        } else if isBackgroundGray(pixelData, offset) {
            grayPixels += 1
        }
    }
    let sampleSize = sampleEnd - sampleStart
    if sampleSize == 0 { continue }
    if slidePixels > sampleSize * 6 / 10 && grayPixels < sampleSize * 2 / 10 {
        topEdge = y
        break
    }
}

// Detect bottom edge by finding last row of slide content before gray UI
// Scan upward from bottom until we find where slide content ends (transition to gray)
var bottomEdge = height - 1
var consecutiveSlideRows = 0
let bottomSampleStart = width / 6
let bottomSampleEnd = width - bottomSampleStart
let minScanY = max(topEdge + 100, height / 2)
for y in stride(from: height - 1, through: minScanY, by: -1) {
    let startOffset = (y * width) * 4
    var slidePixels = 0
    var grayPixels = 0
    var totalPixels = 0
    // Sample middle area
    for x in bottomSampleStart..<bottomSampleEnd {
        let offset = startOffset + (x * 4)
        if isSlidePixel(pixelData, offset) {
            slidePixels += 1
        } else if isBackgroundGray(pixelData, offset) {
            grayPixels += 1
        }
        totalPixels += 1
    }
    if totalPixels == 0 { continue }
    let slidePercent = slidePixels * 100 / totalPixels
    let grayPercent = grayPixels * 100 / totalPixels
    // If we find mostly slide content (non-gray), we're still in the slide
    if slidePercent > 50 && grayPercent < 30 {
        consecutiveSlideRows += 1
        // Found consecutive slide rows - this is the bottom edge
        if consecutiveSlideRows >= 2 {
            bottomEdge = y
            logError("Found bottom edge at row \(y) (last slide row before gray)")
            break
        }
    } else {
        // Hit gray UI or mixed content - reset counter
        consecutiveSlideRows = 0
    }
}

// Use fallback if edges weren't detected properly
if topEdge == 0 {
    logError("Top edge not detected, using fallback 1%")
    topEdge = height * 1 / 100
}
if bottomEdge == height - 1 || bottomEdge <= topEdge {
    logError("Bottom edge not detected, using fallback 11%")
    bottomEdge = height - (height * 11 / 100)
}

// Detect right edge by finding last column of slide content before gray UI
var rightEdge = width - 1
var consecutiveSlideCols = 0
let rightSampleStart = max(topEdge, height / 4)
let rightSampleEnd = min(bottomEdge, height - height / 4)
for x in stride(from: width - 1, through: 0, by: -1) {
    var slidePixels = 0
    var grayPixels = 0
    var totalPixels = 0
    // Sample middle vertical section
    for y in rightSampleStart..<rightSampleEnd {
        let offset = ((y * width) + x) * 4
        if isSlidePixel(pixelData, offset) {
            slidePixels += 1
        } else if isBackgroundGray(pixelData, offset) {
            grayPixels += 1
        }
        totalPixels += 1
    }
    if totalPixels == 0 { continue }
    let slidePercent = slidePixels * 100 / totalPixels
    let grayPercent = grayPixels * 100 / totalPixels
    // If we find mostly slide content (non-gray), we're still in the slide
    if slidePercent > 50 && grayPercent < 30 {
        consecutiveSlideCols += 1
        // Found consecutive slide columns - this is the right edge
        if consecutiveSlideCols >= 2 {
            rightEdge = x
            logError("Found right edge at column \(x) (last slide column before gray)")
            break
        }
    } else {
        // Hit gray UI or mixed content - reset counter
        consecutiveSlideCols = 0
    }
}

// Calculate left edge using 16:9 ratio from detected right edge
var leftEdge: Int
if rightEdge == width - 1 {
    // Fallback: right edge not detected, center horizontally
    logError("Right edge not detected, calculating from 16:9")
    let targetAspectRatio: Double = 16.0 / 9.0
    let slideHeight = bottomEdge - topEdge
    let slideWidth = Int(Double(slideHeight) * targetAspectRatio)
    rightEdge = min(width, width - (width * 3 / 100))
    let leftMargin = max(0, (width - slideWidth) / 2)
    leftEdge = leftMargin
} else {
    // Calculate left edge using 16:9 ratio from detected right edge
    let targetAspectRatio: Double = 16.0 / 9.0
    let slideHeight = bottomEdge - topEdge
    let slideWidth = Int(Double(slideHeight) * targetAspectRatio)
    // Left edge = right edge - slide width
    leftEdge = max(0, rightEdge - slideWidth)
    logError("Calculated left edge at \(leftEdge) using 16:9 from right edge \(rightEdge)")
}

logError("Cropping: x=\(leftEdge) y=\(topEdge) w=\(rightEdge-leftEdge) h=\(bottomEdge-topEdge)")

var finalLeft = leftEdge
var finalTop = topEdge
var finalWidth = rightEdge - leftEdge
var finalHeight = bottomEdge - topEdge

let padding = 3
let cropX = max(0, finalLeft - padding)
let cropY = max(0, finalTop - padding)
let cropWidth = min(width - cropX, finalWidth + padding * 2)
let cropHeight = min(height - cropY, finalHeight + padding * 2)

if cropWidth < 100 || cropHeight < 100 {
    logError("Invalid crop size, aborting")
    exit(1)
}

let cropRect = CGRect(x: cropX, y: cropY, width: cropWidth, height: cropHeight)
guard let croppedCGImage = cgImage.cropping(to: cropRect) else {
    exit(1)
}

let croppedNSImage = NSImage(cgImage: croppedCGImage, size: NSSize(width: cropWidth, height: cropHeight))
guard let tiffData = croppedNSImage.tiffRepresentation,
      let bitmapImage = NSBitmapImageRep(data: tiffData),
      let pngData = bitmapImage.representation(using: .png, properties: [:]) else {
    logError("Failed to create PNG data")
    exit(1)
}

do {
    try pngData.write(to: URL(fileURLWithPath: outputPath))
    exit(0)
} catch {
    logError("Failed to write output file: \(error.localizedDescription)")
    exit(1)
}
SWIFT

chmod +x "$crop_script"

# Run crop detection
cropped_image="/tmp/ppt_slide_cropped_$$.png"
debug_log="/tmp/ppt_debug_$$.log"

# Run swift script and capture output for debugging
if swift "$crop_script" "$temp_image" "$cropped_image" > "$debug_log" 2>&1; then
    # Use cropped image if successful and valid
    if [ -f "$cropped_image" ] && [ -s "$cropped_image" ]; then
        # Verify the cropped image is actually different and reasonable
        original_size=$(stat -f%z "$temp_image" 2>/dev/null || echo 0)
        cropped_size=$(stat -f%z "$cropped_image" 2>/dev/null || echo 0)
        
        # Use cropped if it's valid and smaller (meaning we actually cropped something)
        if [ "$cropped_size" -gt 1000 ] && [ "$cropped_size" -lt "$original_size" ]; then
            mv "$cropped_image" "$temp_image"
        else
            rm -f "$cropped_image"
        fi
    fi
else
    # Edge detection failed - log it
    echo "Swift script failed with exit code $?" >> "$debug_log"
    cat "$debug_log" >> "/tmp/ppt_error_trace.log"
    rm -f "$cropped_image"
fi

rm -f "$crop_script" "$cropped_image"

# Try to use the title from PowerPoint first
slide_title=""
if [[ -n "$ppt_title" ]]; then
    # Clean up the PowerPoint title
    slide_title=$(echo "$ppt_title" | head -n 1 | sed 's/[^a-zA-Z0-9 ]//g' | xargs)
    slide_title="${slide_title:0:50}"
fi

# If we got a good title from PowerPoint, use it; otherwise OCR
if [[ -n "$slide_title" ]] && [[ ${#slide_title} -gt 3 ]]; then
    # We have a good title from PowerPoint, skip OCR
    safe_title="$slide_title"
else
    # Fall back to OCR
    ocr_script="/tmp/ocr_script_$$.swift"
cat > "$ocr_script" <<'SWIFT'
import Cocoa
import Vision

let imagePath = CommandLine.arguments[1]
guard let image = NSImage(contentsOfFile: imagePath) else {
    print("")
    exit(0)
}

guard let fullCGImage = image.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
    print("")
    exit(0)
}

// Crop to top 20% of image where title typically appears (more focused)
let imageWidth = fullCGImage.width
let imageHeight = fullCGImage.height
let cropHeight = Int(Double(imageHeight) * 0.20) // Top 20% - just the title area

let cropRect = CGRect(x: 0, y: imageHeight - cropHeight, width: imageWidth, height: cropHeight)
guard let croppedImage = fullCGImage.cropping(to: cropRect) else {
    print("")
    exit(0)
}

var result = ""
let semaphore = DispatchSemaphore(value: 0)

let request = VNRecognizeTextRequest { request, error in
    defer { semaphore.signal() }
    
    guard let observations = request.results as? [VNRecognizedTextObservation] else {
        return
    }
    
    // Find the topmost text (highest on screen = highest Y value)
    // Prioritize longer text (titles are usually longer than UI elements)
    var topmostText = ""
    var maxY: CGFloat = -1
    var bestScore: CGFloat = -1
    
    for observation in observations {
        guard let topCandidate = observation.topCandidates(1).first else { continue }
        let text = topCandidate.string.trimmingCharacters(in: .whitespaces)
        if text.isEmpty { continue }
        
        // Filter out common UI text patterns
        let lowerText = text.lowercased()
        if lowerText.contains("search") || 
           lowerText.contains("cmd") || 
           lowerText.contains("ctrl") ||
           lowerText.contains("click to add") ||
           lowerText.contains("presenter notes") ||
           (lowerText.contains("notes") && text.count < 20) ||
           text.count < 4 {
            continue
        }
        
        // Get the bounding box - higher Y means higher on screen
        let boundingBox = observation.boundingBox
        let y = boundingBox.origin.y + boundingBox.height
        
        // Score: prioritize both height (Y) and length (longer = more likely to be title)
        let score = y + (CGFloat(text.count) * 0.01)
        
        if score > bestScore {
            bestScore = score
            maxY = y
            topmostText = text
        }
    }
    
    // If we found text, use it; otherwise try first observation
    if !topmostText.isEmpty {
        result = topmostText
    } else if let firstObs = observations.first,
              let topCandidate = firstObs.topCandidates(1).first {
        let candidateText = topCandidate.string.trimmingCharacters(in: .whitespaces)
        // Filter UI text
        let lowerText = candidateText.lowercased()
        if !lowerText.contains("search") && 
           !lowerText.contains("cmd") && 
           !lowerText.contains("ctrl") &&
           !lowerText.contains("click to add") &&
           !lowerText.contains("presenter notes") &&
           !(lowerText.contains("notes") && candidateText.count < 15) &&
           candidateText.count >= 3 {
            result = candidateText
        }
    }
    
    // Take only first line and clean up
    if !result.isEmpty {
        let lines = result.components(separatedBy: .newlines)
        result = lines.first?.trimmingCharacters(in: .whitespaces) ?? ""
        // Remove common OCR artifacts
        result = result.replacingOccurrences(of: "|", with: "I")
        result = result.replacingOccurrences(of: "0", with: "O")
    }
}

request.recognitionLevel = .accurate
request.usesLanguageCorrection = false

let handler = VNImageRequestHandler(cgImage: croppedImage, options: [:])
do {
    try handler.perform([request])
    _ = semaphore.wait(timeout: .now() + 5.0)
    print(result)
} catch {
    print("")
}
SWIFT

    # Run OCR
    slide_title=$(swift "$ocr_script" "$temp_image" 2>/dev/null | head -n 1 | xargs)
    rm -f "$ocr_script"

    # Clean up slide title
    safe_title=$(echo "$slide_title" | sed 's/[^a-zA-Z0-9 ]//g' | xargs)
    safe_title="${safe_title:0:50}"
    
    # If OCR failed, use slide number
    if [[ -z "$safe_title" ]] || [[ ${#safe_title} -lt 3 ]]; then
        if [[ -n "$slide_num" ]] && [[ "$slide_num" =~ ^[0-9]+$ ]]; then
            safe_title="Slide $slide_num"
        else
            safe_title="Slide"
        fi
    fi
fi

# Create filename
filename="${course_code} ${safe_title}.png"
target_dir="$HOME/Downloads"
final_path="${target_dir}/${filename}"

# Handle duplicates
counter=2
while [ -f "$final_path" ]; do
    filename="${course_code} ${safe_title} ${counter}.png"
    final_path="${target_dir}/${filename}"
    ((counter++))
done

# Move temp file to final location
mv "$temp_image" "$final_path"

if [ $? -eq 0 ]; then
    echo "Saved: $filename"
else
    echo "ERROR: Failed to save file"
    exit 1
fi
