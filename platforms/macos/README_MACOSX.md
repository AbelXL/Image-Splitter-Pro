# Technical Details

This project uses `pyproject.toml` for dependency management.

# Compilation

Run the following command from the root directory of the project to compile the application.

pyinstaller --clean --onedir --windowed --name ImageSplitterPro --hidden-import=ttkbootstrap --hidden-import=tkinterdnd2 --add-data "platforms/macos/icon.icns:." --add-data "src/image_splitter_pro/profile_editor.py:." --icon "platforms/macos/icon.icns" src/image_splitter_pro/main.py

# Running the Application

Once completed, the executable will be located in the dist/ImageSplitterPro folder.