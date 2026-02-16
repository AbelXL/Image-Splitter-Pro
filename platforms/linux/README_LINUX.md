# Technical Details

This project uses `pyproject.toml` for dependency management.

## Compilation

Run the following command from the root directory of the project to compile the application.

pyinstaller --clean --onedir --windowed \
  --name ImageSplitterPro \
  --hidden-import=ttkbootstrap \
  --hidden-import=PIL._tkinter_finder \
  --hidden-import=PIL._imagingtk \
  --hidden-import=PIL.ImageTk \
  --hidden-import=PIL.Image \
  --hidden-import=tkinterdnd2 \
  --add-data "platforms/linux/icon.ico:." \
  --add-data "src/image_splitter_pro/profile_editor.py:." \
  --icon "platforms/linux/icon.ico" \
  src/image_splitter_pro/main.py

## Running the Application

Once completed, the executable will be located in the dist/ImageSplitterPro folder.

## Create Desktop Launcher

Use the create-desktop-entry.sh script to create shortcuts on your desktop, in your application folder, and in your system menu.

### Copy the script into the directory with the application
cp platforms/linux/create-desktop-entry.sh dist/ImageSplitterPro/

### Make the script executable
chmod +x dist/ImageSplitterPro/create-desktop-entry.sh

### Run the script
./dist/ImageSplitterPro/create-desktop-entry.sh


