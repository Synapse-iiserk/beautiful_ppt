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
3. Update the image src and alt attributes in `synapse_platform_presentation.html` around line 1860:
   ```html
   <img src="your_logo_1.png" alt="Sequoia Capital logo" style="height: 50px; opacity: 0.7; filter: grayscale(100%);" />
   ```

**Important:** When replacing logos, always update the `alt` attribute with descriptive text (e.g., "Sequoia Capital logo", "Andreessen Horowitz logo") for accessibility and screen reader users.

## Google Drive Links (from issue)
The original logo images were referenced with these identifiers:
1. UOu5loURrOcmP7eqh
2. CFin470jnYhOSzjGH
3. tifOc9nWV4buVAU1M
4. eHtnDocpXirneQV90
5. nCQs6rPddjon8Oxkl

**Note:** These appear to be shortened Google Drive identifiers. You may need to obtain the actual Google Drive share links or download the images directly from the source provided in the issue.

## Styling
The logos are styled with:
- Height: 50px (maintains aspect ratio)
- Opacity: 0.7 (70% opacity)
- Filter: grayscale(100%) (black and white)
- Gap between logos: var(--space-lg)
- Responsive flex layout with wrapping

You can adjust these styles in the HTML if needed.
