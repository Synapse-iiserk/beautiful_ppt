# VC and Company Recognition Logos

## Overview
The last slide (Slide 14: Contact) of `synapse_platform_presentation.html` now includes a "Recognised by the following VCs and companies" section with placeholder logos.

## Replacing Placeholder Logos

The current implementation uses placeholder SVG files (vc_logo_1.svg through vc_logo_5.svg). To replace them with actual VC/company logos:

### Option 1: Replace the SVG files
1. Download the actual logo images from the provided Google Drive links
2. Convert them to SVG format (recommended) or PNG/JPG
3. Name them as:
   - `vc_logo_1.svg` (or .png/.jpg)
   - `vc_logo_2.svg`
   - `vc_logo_3.svg`
   - `vc_logo_4.svg`
   - `vc_logo_5.svg`
4. Replace the placeholder files in the repository root

### Option 2: Update the HTML directly
1. Download the actual logo images
2. Save them with meaningful names (e.g., `sequoia_capital.png`, `andreessen_horowitz.png`)
3. Update the image src attributes in `synapse_platform_presentation.html` around line 1860:
   ```html
   <img src="your_logo_1.png" alt="Company Name 1" style="height: 50px; opacity: 0.7; filter: grayscale(100%);" />
   ```

## Google Drive Links (from issue)
The original logo images can be found at:
1. https://share.google/UOu5loURrOcmP7eqh
2. https://share.google/CFin470jnYhOSzjGH
3. https://share.google/tifOc9nWV4buVAU1M
4. https://share.google/eHtnDocpXirneQV90
5. https://share.google/nCQs6rPddjon8Oxkl

**Note:** These links may need to be converted to full Google Drive URLs for downloading.

## Styling
The logos are styled with:
- Height: 50px (maintains aspect ratio)
- Opacity: 0.7 (70% opacity)
- Filter: grayscale(100%) (black and white)
- Gap between logos: var(--space-lg)
- Responsive flex layout with wrapping

You can adjust these styles in the HTML if needed.
