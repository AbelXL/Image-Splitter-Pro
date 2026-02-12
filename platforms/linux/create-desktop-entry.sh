#!/bin/bash

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Define the executable name
EXECUTABLE_NAME="ImageSplitterPro"
EXECUTABLE_PATH="$SCRIPT_DIR/$EXECUTABLE_NAME"

# Define the desktop entry filename
DESKTOP_ENTRY_NAME="ImageSplitterPro.desktop"
DESKTOP_ENTRY_PATH="$SCRIPT_DIR/$DESKTOP_ENTRY_NAME"

# Check if the executable exists
if [ ! -f "$EXECUTABLE_PATH" ]; then
    echo "Error: Executable not found at $EXECUTABLE_PATH"
    exit 1
fi

# Look for icon file (check for .ico or .png)
ICON_PATH=""
if [ -f "$SCRIPT_DIR/_internal/icon.ico" ]; then
    ICON_PATH="$SCRIPT_DIR/_internal/icon.ico"
elif [ -f "$SCRIPT_DIR/icon.ico" ]; then
    ICON_PATH="$SCRIPT_DIR/icon.ico"
fi

# Create the .desktop file
# Note: For .desktop files with paths containing spaces, we use sh -c to properly handle the path
cat > "$DESKTOP_ENTRY_PATH" << 'DESKTOPEOF'
[Desktop Entry]
Version=1.0
Type=Application
Name=Image Splitter Pro
Comment=Image processing and cropping tool
DESKTOPEOF

# Add Exec line with properly escaped path
echo "Exec=sh -c 'cd \"\$1\" && ./ImageSplitterPro' sh \"$SCRIPT_DIR\"" >> "$DESKTOP_ENTRY_PATH"

# Continue with the rest of the desktop entry
cat >> "$DESKTOP_ENTRY_PATH" << 'DESKTOPEOF'
Terminal=false
StartupNotify=false
DESKTOPEOF

# Add icon if found
if [ -n "$ICON_PATH" ]; then
    echo "Icon=$ICON_PATH" >> "$DESKTOP_ENTRY_PATH"
fi

# Add categories
echo "Categories=Graphics;Utility;" >> "$DESKTOP_ENTRY_PATH"

# Make the .desktop file executable
chmod +x "$DESKTOP_ENTRY_PATH"

# Make the executable file executable (if not already)
chmod +x "$EXECUTABLE_PATH"

echo "Desktop entry created successfully at:"
echo "$DESKTOP_ENTRY_PATH"
echo ""

# Copy to ~/.local/share/applications/ for applications menu
APPS_DIR="$HOME/.local/share/applications"
if [ ! -d "$APPS_DIR" ]; then
    mkdir -p "$APPS_DIR"
    echo "Created directory: $APPS_DIR"
fi

cp "$DESKTOP_ENTRY_PATH" "$APPS_DIR/"
chmod +x "$APPS_DIR/$DESKTOP_ENTRY_NAME"
echo "✓ Installed to applications menu: $APPS_DIR/$DESKTOP_ENTRY_NAME"

# Copy to ~/Desktop/ for desktop shortcut
DESKTOP_DIR="$HOME/Desktop"
if [ -d "$DESKTOP_DIR" ]; then
    cp "$DESKTOP_ENTRY_PATH" "$DESKTOP_DIR/"
    chmod +x "$DESKTOP_DIR/$DESKTOP_ENTRY_NAME"

    # Mark as trusted (for Ubuntu/GNOME to allow launching)
    if command -v gio &> /dev/null; then
        gio set "$DESKTOP_DIR/$DESKTOP_ENTRY_NAME" metadata::trusted true 2>/dev/null || true
    fi

    echo "✓ Installed to desktop: $DESKTOP_DIR/$DESKTOP_ENTRY_NAME"
else
    echo "⚠ Desktop directory not found, skipping desktop shortcut"
fi

echo ""
echo "Installation complete!"
echo "You can now:"
echo "  • Double-click '$DESKTOP_ENTRY_NAME' on your desktop to launch Image Splitter Pro"
echo "  • Find 'Image Splitter Pro' in your applications menu"
echo "  • Double-click '$DESKTOP_ENTRY_NAME' in this directory"

