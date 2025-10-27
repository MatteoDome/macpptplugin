# PowerPoint Shortcuts - Setup Complete! 🎉

## What I've Done

✅ **Uploaded all add-in files** to your `macpptplugin` repository  
✅ **Updated manifest.json** with correct GitHub Pages URLs  
✅ **Created placeholder icons** (you can replace these later)  
✅ **Committed and pushed** all files to GitHub  

## Next Steps (You need to do these)

### 1. Enable GitHub Pages
1. Go to https://github.com/MatteoDome/macpptplugin
2. Click **Settings** tab
3. Scroll down to **Pages** in the left sidebar
4. Under "Source", select **"Deploy from a branch"**
5. Choose **"main"** branch and **"/ (root)"**
6. Click **Save**

GitHub will show you the URL: `https://matteodome.github.io/macpptplugin`

### 2. Wait for Deployment (2-3 minutes)
GitHub Pages takes a few minutes to deploy. You'll see a green checkmark when it's ready.

### 3. Test the Deployment
Visit these URLs to make sure they work:
- Main site: https://matteodome.github.io/macpptplugin
- Task pane: https://matteodome.github.io/macpptplugin/src/taskpane/taskpane.html
- Commands: https://matteodome.github.io/macpptplugin/src/commands/commands.html

### 4. Install in PowerPoint
1. Download the `manifest.xml` file from your repository
2. Open PowerPoint for Mac
3. Go to **Insert** → **My Add-ins** → **Upload My Add-in**
4. Select the `manifest.xml` file
5. The add-in should appear in your ribbon!

## Your Add-in URLs

Once GitHub Pages is enabled, your add-in will be available at:
- **Repository**: https://github.com/MatteoDome/macpptplugin
- **Live site**: https://matteodome.github.io/macpptplugin
- **Manifest**: https://matteodome.github.io/macpptplugin/manifest.xml

## Features Available

### Keyboard Shortcuts
- `Ctrl+Alt+T` - Paste unformatted text
- `Shift+Alt+Z` - Convert text to autoshape
- `Shift+Alt+E` - Make same width
- `Shift+Alt+H` - Make same height
- `Ctrl+Alt+C` - Align center
- `Ctrl+Alt+L` - Align left
- `Ctrl+Alt+R` - Align right
- `Alt+Shift+D` - Distribute horizontally
- `Ctrl+Alt+F` - Insert footnote
- `Ctrl+Alt+S` - Insert sticker
- And many more!

### Task Pane
- Visual interface for all shortcuts
- Professional color palette
- Element insertion tools
- Help and documentation

## Troubleshooting

### If GitHub Pages isn't working:
1. Make sure the repository is public
2. Wait 5-10 minutes after enabling Pages
3. Check that files are in the main branch

### If PowerPoint won't load the add-in:
1. Verify the URLs work in your browser
2. Check that manifest.xml is valid
3. Make sure you downloaded the latest manifest.xml

### If you get icon errors:
The placeholder icons I created are empty files. You can:
1. Ignore the errors (add-in will work fine)
2. Create real PNG icons and upload them
3. Remove icon references from manifest.json

## Creating Real Icons (Optional)

If you want proper icons:
1. Create PNG files: 16x16, 32x32 pixels
2. Use professional blue (#0078D4) background
3. Add white "P" or "PS" text
4. Upload to the `assets` folder
5. Name them: `icon-16.png`, `icon-32.png`, `icon-outline.png`, `icon-color.png`

## Updating the Add-in

To update later:
1. Edit files in your repository
2. Commit and push changes
3. GitHub Pages updates automatically
4. PowerPoint will use the new version

## Support

If you need help:
1. Check the URLs work in your browser
2. Validate manifest.xml using Office Add-in validator
3. Look at browser console for errors

**Your add-in is ready to go! Just enable GitHub Pages and install in PowerPoint.** 🚀

