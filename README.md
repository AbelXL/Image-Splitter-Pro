## 📸 Image Splitter Pro

Take a single photo, split it into several images, and upload them to your favorite platforms.

Image Splitter Pro is a free, open-source desktop application built with privacy at its core. Available for Windows, macOS, and Linux, it ensures your photos never leave your computer. Unlike web-based tools, all processing happens locally—meaning no cloud uploads, no data tracking, and no privacy compromises.

## 📱 Social Media: One master image, many platform-ready crops.

• 1:1 – Instagram Posts & Profile Pics

• 9:16 – TikTok, Reels & Stories

• 4:5 – High-Engagement Portrait Feed

• 1.91:1 – Professional Facebook & LinkedIn Ads

• Custom – Set your own dimensions for any site


## 🛍️ eCommerce Optimized (eBay, Etsy, Poshmark)

Drop in one high-resolution shirt photo and generate a suite of detailed crops:

• Collar (Brand tags and stitching)

• Sleeve (Cuffs and texture)

• Hem (Finish and quality)

## 🚀 Support the Project - [www.abelxl.com](https://www.abelxl.com/)

## ✨ Key Features

• Batch Processing: Create and apply custom cropping rules to hundreds of images in seconds.

• Auto-Folder Export: Automatically organizes your projects into labeled directories. Keeps your workspace clean and your images ready for instant upload to eBay or social media.

• Supported Formats: Image Splitter Pro handles all major image types, including JPG, PNG, WebP, BMP, GIF, TIFF, and HEIC/HEIF.

• Format Consistency: To preserve your workflow, all cropped images maintain their original format (e.g., a PNG stays a PNG). The only exceptions are HEIC/HEIF files, which are automatically converted to JPG for maximum compatibility with eBay and social media platforms.


## 📐 Supported Aspect Ratios

• Image Splitter Pro offers a comprehensive list of presets to ensure your crops fit most platforms. You can choose from: 1:1, 3:4, 4:3, 4:5, 16:9, 9:16, 3:2, 1.91:1

• Custom: Set your own specific proportions.

• None: Maintain the original aspect ratio of your source image.


## 🚀 How It Works

Getting professional crops is a simple three-step process:

1. Load Your Images

Drag and drop your photos directly into the app, or place them in the Source folder (found under the File menu).

2. Create Your Cropping Profile

    Go to Edit > Create Profile to open the Profile Editor.

    Select your image and choose an Aspect Ratio.

    Use your mouse to resize and position the crop exactly where you want it.

    Add Multiple Rules: Click Create New Cropping Rule to create additional crops from the same image (e.g., one for the collar, one for the sleeve) or select a different image.

    Name your profile and click Save & Close.

    Your profile is now saved in the ‘config’ folder as Your_Name.profile

3. Process and Export

    Select your new profile from the list and click the Crop button.

    Review your results instantly in the Image Preview panel.

    Click Move to automatically create a dedicated project folder in your Destination directory and move all finished images there.

    
[!WARNING] ⚠️ Data Safety: The Move function transfers images from the Source to the Destination folder. Any other files remaining in the Source folder may be moved or deleted during this process. Always back up your original photos before processing.



## 💡 Advanced Workflow: Understanding "Positions"

Image Splitter Pro uses Position Logic to turn your cropping rules into reusable templates, called Profiles. When you select an image in the Profile Editor, it is assigned a label (e.g., Position 1, Position 2).


## 🏗️ How Profiles Work: The app remembers the exact pixel coordinates of your crops for each Position. This means:

• You create a profile called "Men's Shirts" with three crops (Collar, Sleeve, Hem) on the image in Position 1.

• Tomorrow, you drag in a new shirt photo.

• You simply select the "Men's Shirts" profile and click Crop.

• The app automatically applies those same three crops to the new photo instantly.

• 💥Apply to Remaining Images Feature

The Apply to Remaining Images option in Image Splitter Pro allows you to apply a single cropping rule to multiple images at once, without creating a separate rule for each position.


## ⚙️ How It Works

When you enable "Apply to Remaining Images" on a rule:

• Identifies Available Positions: The app looks at all images in your Source Folder that don't already have specific rules assigned to them.

• Applies the Rule Automatically: Your cropping rule (crop dimensions, output size, etc.) is applied to every remaining image that doesn't have an explicit rule.

• Respects Rule Order: If you have multiple rules with "Apply to Remaining Images" enabled, they are applied in the order they appear in your profile file.


Example Use Case

Let's say you have 10 images and want to:

1. Crop image #1 as a profile picture (800x1000)

2. Crop image #2 as a header banner (1000x1500)

3. Crop images #3-10 all the same way (500x500)

	a. Without "Apply to Remaining Images". You would need to create 8 separate rules for images #3-10.

Steps with "Apply to Remaining Images":

1. Create Rule 1 for Position 1 (800x1000).

2. Create Rule 2 for Position 2 (1000x15000).

3. Create Rule 3 with "Apply to Remaining Images" checked (500x500).

	a. Rule 3 automatically applies to all remaining images (#3-10), saving you time and effort.

Important Notes

• Position-Specific Rules Take Priority: If an image has a specific position rule, it won't be affected by "Apply to Remaining Images" rules.

• Multiple "Apply to Remaining" Rules: If you have more than one rule with this option enabled, they work together to fill any gaps. The first rule in your profile handles unfilled positions first, then the second rule, and so on.

• Order Matters: Rules are applied in the order they appear in your profile, so position your "Apply to Remaining Images" rules accordingly.

This feature is perfect for batch processing images where most follow the same pattern, with only a few exceptions needing custom treatment.



## 📏 Best Practices for Perfect Results

To get consistent, professional results when reusing profiles, follow these guidelines:

• Uniform Image Size: Ensure all source images for a specific profile are the same dimensions (e.g., all are 4000x4000px). Since crops are pixel-based, mixing image sizes will cause your crops to shift.

• Consistent Framing: Keep your subject (e.g., the shirt) in the same spot in the frame for every photo. If the shirt moves, the fixed crop coordinates won't align.

• High Resolution: Always use high-resolution source images. Because you are cropping into small details, starting with a large file ensures your final image close-ups stay sharp.

• Sorting Logic: Images are sorted from Oldest to Newest. If the order looks incorrect, ensure your file explorer is set to sort by Date Modified rather than by Name. Image Splitter Pro defaults to chronological order.


## 🛠️ Built With & Third-Party Credits

This application is powered by the following open-source software. We are grateful to the developers who maintain these tools:

• Python: Created by the Python Software Foundation.

• Pillow (PIL): Used for high-performance image manipulation. [License: HPND]

• ttkbootstrap: Used for the modern user interface. [License: MIT]

• Tkinter: Python's standard GUI framework.

• tkinterdnd2: Drag-and-drop functionality [MIT License]

• send2trash: Safe file deletion [BSD License]

• pillow-heif: HEIC format support [BSD License]


## 🛠️ Installation & Compilation

Platform Specific Instructions

For detailed compilation instructions, please see the specific guides:

Windows: Instructions in [`platforms/windows/README_WINDOWS.md`](platforms/windows/README_WINDOWS.md)  
macOS: Instructions in [`platforms/macos/README_OSX.md`](platforms/macos/README_MACOSX.md)  
Linux: Instructions in [`platforms/linux/README_LINUX.md`](platforms/linux/README_LINUX.md)   

## ⚖️ Legal Information

Image Splitter Pro bundles several third-party libraries under their respective open-source licenses (MIT, BSD, and PSF). By using this software, you agree to the terms of these original licenses. The full text for these licenses can be found in the source code repositories of the respective projects.


## ✍️ Author

Abel Aramburo | [@Abel_XL](https://x.com/Abel_XL) | [www.abelxl.com](https://www.abelxl.com/)

This project was developed with the assistance of AI logic-modeling to ensure high-performance image handling and a modern user experience.


## 📄 License

This project is licensed under the MIT License. This means you are free to use, modify, and distribute the software, provided that the original copyright notice and this permission notice are included in all copies or substantial portions of the software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.

Copyright (c) 2026 Abel Aramburo

We hope this tool saves you countless hours and makes your workflow feel effortless.

Happy cropping! 📸✨
