#!/bin/bash

# Script to generate a soft link creation script from an m3u playlist
# Usage: ./generate_playlist_links.sh <path_to_m3u> <target_directory>

# Check if correct number of arguments provided
if [ $# -ne 2 ]; then
    echo "Usage: $0 <m3u_file> <target_directory>"
    exit 1
fi

M3U_FILE="$1"
TARGET_DIR="$2"

# Check if m3u file exists
if [ ! -f "$M3U_FILE" ]; then
    echo "Error: M3U file '$M3U_FILE' not found"
    exit 1
fi

# Extract the m3u filename without path
M3U_BASENAME=$(basename "$M3U_FILE")

# Get the directory where the m3u file is located
M3U_DIR=$(dirname "$M3U_FILE")

# Create output script name
OUTPUT_SCRIPT="${M3U_DIR}/linkfiles_${M3U_BASENAME}.sh"

# Start creating the output script
cat > "$OUTPUT_SCRIPT" << 'HEADER'
#!/bin/bash

# Auto-generated script to create soft links for playlist files
# Target directory: TARGET_DIR_PLACEHOLDER

TARGET_DIR="TARGET_DIR_PLACEHOLDER"
M3U_OUTPUT="M3U_OUTPUT_PLACEHOLDER"

# Create target directory if it doesn't exist
mkdir -p "$TARGET_DIR"

# Create new m3u playlist file
echo "#EXTM3U" > "$M3U_OUTPUT"

echo "Creating soft links in $TARGET_DIR..."

HEADER

# Replace the placeholders with actual values
sed -i "s|TARGET_DIR_PLACEHOLDER|$TARGET_DIR|g" "$OUTPUT_SCRIPT"
sed -i "s|M3U_OUTPUT_PLACEHOLDER|$TARGET_DIR/$M3U_BASENAME|g" "$OUTPUT_SCRIPT"

# Process each line in the m3u file
while IFS= read -r line || [ -n "$line" ]; do
    # Skip empty lines and comments (lines starting with #)
    if [[ -z "$line" || "$line" =~ ^[[:space:]]*# ]]; then
        continue
    fi
    
    # Convert Windows path separators to Linux
    linux_path=$(echo "$line" | sed 's/\\/\//g')
    
    # Remove any carriage returns (in case of Windows line endings)
    linux_path=$(echo "$linux_path" | tr -d '\r')
    
    # Prepend the absolute path
    full_path="/share/Music/${linux_path}"
    
    # Get the relative directory structure (everything except the filename)
    relative_dir=$(dirname "$linux_path")
    filename=$(basename "$linux_path")
    
    # Escape special characters for shell
    # We need to escape: space, (, ), [, ], {, }, ', ", `, $, &, ;, |, <, >, !, *, ?, ~, \
    escaped_path=$(printf '%q' "$full_path")
    escaped_relative_dir=$(printf '%q' "$relative_dir")
    escaped_filename=$(printf '%q' "$filename")
    
    # Build the absolute path for the m3u file
    if [ -n "$relative_dir" ] && [ "$relative_dir" != "." ]; then
        absolute_m3u_path="\$TARGET_DIR/$relative_dir/$filename"
    else
        absolute_m3u_path="\$TARGET_DIR/$filename"
    fi
    
    # Add commands to the output script to create directory structure and link
    echo "mkdir -p \"\$TARGET_DIR\"/$escaped_relative_dir" >> "$OUTPUT_SCRIPT"
    echo "ln -sf $escaped_path \"\$TARGET_DIR\"/$escaped_relative_dir/$escaped_filename" >> "$OUTPUT_SCRIPT"
    echo "echo \"$absolute_m3u_path\" >> \"\$M3U_OUTPUT\"" >> "$OUTPUT_SCRIPT"
    
done < "$M3U_FILE"

# Add completion message
echo "" >> "$OUTPUT_SCRIPT"
echo 'echo "Soft link creation complete!"' >> "$OUTPUT_SCRIPT"
echo 'echo "Playlist created: $M3U_OUTPUT"' >> "$OUTPUT_SCRIPT"

# Make the output script executable
chmod +x "$OUTPUT_SCRIPT"

echo "Generated script: $OUTPUT_SCRIPT"
echo "Run it with: $OUTPUT_SCRIPT"
