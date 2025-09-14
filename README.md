# Automated-PowerPoint-Report-Generation

This project automates the creation of PowerPoint slides for social media influencers from an Excel template. It uses Excel VBA macros and integrates with PowerPoint to generate fully formatted slides with dynamic data, clickable icons, and images.

🔹 Project Structure
File	Description
Tim_social_links - Tamplate.xlsm	Excel template containing influencer data. Columns include Name, Username, Followers, Date, Instagram, Facebook, TikTok, Snapchat, Images, and PCloud Links.
Orignal Tamp to generate slids.pptx	PowerPoint template used to create slides. Slide 10 is used as the main template.
VBA Code For Creating Report Slides.txt	VBA macro to generate PowerPoint slides dynamically from Excel data.
CoverageIcon.txt	VBA macro to hyperlink coverage icons in slides to PCloud links from Excel.
Socialmediaicons.txt	VBA macro to hyperlink social media icons in each slide based on Excel data.
ReplaceImagesFomFolders.txt	VBA macro to replace placeholder images in slides with actual coverage images.
generated_presentation_finll.pptm	Output PowerPoint file after running the macros (contains all slides, hyperlinks, and images).
🔹 Workflow
Step 1: Generate Initial Slides

Open Tim_social_links - Tamplate.xlsm in Excel.

Enable macros.

Run the macro from VBA Code For Creating Report Slides.txt.

This macro will:

Open Orignal Tamp to generate slids.pptx.

Duplicate Slide 10 for each influencer row in the Excel sheet.

Replace placeholders with influencer data:

Name

Username

Followers

Date

Social Media handles

Format the "Title" and text fields.

Save the generated presentation as generated_presentation_finll.pptm.

Result: A PowerPoint file with slides populated with influencer data and some hyperlinks (if included).

Step 2: Hyperlink Coverage Icons

Open generated_presentation_finll.pptm.

Run the macro from CoverageIcon.txt.

This macro will:

Link the coverage icon on each slide to the corresponding PCloud Link in the Excel sheet.

Result: Coverage icons are now clickable hyperlinks pointing to PCloud resources.

⚠️ For security/testing, real links may be removed in some cases.

Step 3: Hyperlink Social Media Icons

Run the macro from Socialmediaicons.txt.

The macro will:

Add hyperlinks to Instagram, Facebook, TikTok, and Snapchat icons in each slide based on Excel data.

Result: Social media icons are now active hyperlinks.

Step 4: Replace Images in Slides

Run the macro from ReplaceImagesFomFolders.txt.

This macro will:

Replace placeholder images in each slide with screenshots or coverage images from a designated folder.

Fill three images per slide.

Result: Every slide now contains three images representing coverage or campaign visuals. The last 3 slides showcase the final visual coverage results.

🔹 Final Output

generated_presentation_finll.pptm contains:

All dynamically generated slides from Excel data.

Clickable coverage icons linked to PCloud.

Clickable social media icons.

Inserted images for coverage screenshots.

The presentation is ready for final review or distribution.

🔹 Notes

Dependencies: Microsoft Excel and PowerPoint (with macro support).

Font: The template uses "League Spartan". Ensure it is installed.

Security: PCloud and social media links are sensitive; adjust macros if using dummy or test data.

Automation: Fully automated report generation, requiring minimal manual edits.
