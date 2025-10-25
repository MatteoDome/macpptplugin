/*
 * PowerPoint Shortcuts - Commands
 * Keyboard shortcut implementations
 */

// Initialize Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("PowerPoint Shortcuts loaded");
        registerKeyboardShortcuts();
    }
});

// Register all keyboard shortcuts
function registerKeyboardShortcuts() {
    // Text and Object Manipulation
    Office.actions.associate("PasteUnformatted", pasteUnformattedText);
    Office.actions.associate("TextToAutoshape", textToAutoshape);
    Office.actions.associate("SplitJoinTextboxes", splitJoinTextboxes);
    
    // Alignment and Layout
    Office.actions.associate("MakeSameWidth", makeSameWidth);
    Office.actions.associate("AlignCenter", alignCenter);
    Office.actions.associate("AlignLeft", alignLeft);
    Office.actions.associate("AlignRight", alignRight);
    Office.actions.associate("AlignMiddle", alignMiddle);
    Office.actions.associate("DistributeHorizontally", distributeHorizontally);
    Office.actions.associate("DistributeVertically", distributeVertically);
    
    // Professional Elements
    Office.actions.associate("InsertFootnote", insertFootnote);
    Office.actions.associate("InsertLegend", insertLegend);
    Office.actions.associate("InsertSticker", insertSticker);
    Office.actions.associate("CycleAccentColors", cycleAccentColors);
    
    // Utility Functions
    Office.actions.associate("GreenPrint", greenPrint);
    Office.actions.associate("ShowQuickKeys", showQuickKeys);
    Office.actions.associate("ResetElements", resetElements);
    Office.actions.associate("ShowTaskpane", showTaskpane);
}

// Text and Object Manipulation Functions
async function pasteUnformattedText() {
    try {
        await PowerPoint.run(async (context) => {
            const selection = context.presentation.getSelectedTextRange();
            
            // Get clipboard content (simplified - in real implementation would need more complex clipboard handling)
            const clipboardText = await navigator.clipboard.readText();
            
            if (selection && clipboardText) {
                selection.text = clipboardText;
                // Remove formatting by setting to default
                selection.font.name = "Calibri";
                selection.font.size = 18;
                selection.font.bold = false;
                selection.font.italic = false;
                selection.font.underline = PowerPoint.UnderlineType.none;
            }
            
            await context.sync();
        });
        showNotification("Text pasted without formatting");
        return;
    } catch (error) {
        console.error("Error pasting unformatted text:", error);
        showNotification("Error: Could not paste unformatted text", "error");
        return error.code;
    }
}

async function textToAutoshape() {
    try {
        await PowerPoint.run(async (context) => {
            const selection = context.presentation.getSelectedTextRange();
            
            if (selection) {
                const selectedText = selection.text;
                const slide = context.presentation.getSelectedSlides().getItemAt(0);
                
                // Create a rounded rectangle shape
                const shape = slide.shapes.addGeometricShape(
                    PowerPoint.GeometricShapeType.roundRectangle,
                    {
                        left: 100,
                        top: 100,
                        width: 200,
                        height: 50
                    }
                );
                
                // Set the text
                shape.textFrame.textRange.text = selectedText;
                
                // Apply Professional styling
                shape.fill.setSolidColor("#00A651"); // Professional green
                shape.textFrame.textRange.font.color = "#FFFFFF";
                shape.textFrame.textRange.font.name = "Calibri";
                shape.textFrame.textRange.font.size = 14;
                shape.textFrame.textRange.font.bold = true;
                
                // Remove original text if it was in a text box
                selection.text = "";
                
                await context.sync();
            }
        });
        showNotification("Text converted to autoshape");
        return;
    } catch (error) {
        console.error("Error converting text to autoshape:", error);
        showNotification("Error: Could not convert text to autoshape", "error");
        return error.code;
    }
}

async function splitJoinTextboxes() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            // Get selected shapes
            const selectedShapes = shapes.items.filter(shape => 
                shape.type === PowerPoint.ShapeType.geometricShape && 
                shape.textFrame.hasText
            );
            
            if (selectedShapes.length === 1) {
                // Split textbox - create two textboxes from one
                const originalShape = selectedShapes[0];
                const text = originalShape.textFrame.textRange.text;
                const midPoint = Math.floor(text.length / 2);
                
                const firstHalf = text.substring(0, midPoint);
                const secondHalf = text.substring(midPoint);
                
                // Update original with first half
                originalShape.textFrame.textRange.text = firstHalf;
                
                // Create new shape with second half
                const newShape = slide.shapes.addTextBox(
                    firstHalf,
                    {
                        left: originalShape.left + originalShape.width + 10,
                        top: originalShape.top,
                        width: originalShape.width,
                        height: originalShape.height
                    }
                );
                newShape.textFrame.textRange.text = secondHalf;
                
            } else if (selectedShapes.length === 2) {
                // Join textboxes - combine text from two textboxes
                const shape1 = selectedShapes[0];
                const shape2 = selectedShapes[1];
                
                const combinedText = shape1.textFrame.textRange.text + " " + shape2.textFrame.textRange.text;
                shape1.textFrame.textRange.text = combinedText;
                
                // Delete the second shape
                shape2.delete();
            }
            
            await context.sync();
        });
        showNotification("Textboxes split/joined successfully");
        return;
    } catch (error) {
        console.error("Error splitting/joining textboxes:", error);
        showNotification("Error: Could not split/join textboxes", "error");
        return error.code;
    }
}

// Alignment Functions
async function makeSameWidth() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id); // Get all shapes for demo
            
            if (selectedShapes.length > 1) {
                const referenceWidth = selectedShapes[0].width;
                
                for (let i = 1; i < selectedShapes.length; i++) {
                    selectedShapes[i].width = referenceWidth;
                }
                
                await context.sync();
            }
        });
        showNotification("Objects resized to same width");
        return;
    } catch (error) {
        console.error("Error making same width:", error);
        showNotification("Error: Could not make objects same width", "error");
        return error.code;
    }
}

async function alignCenter() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            
            if (selectedShapes.length > 0) {
                const slideWidth = 720; // Standard slide width in points
                
                selectedShapes.forEach(shape => {
                    shape.left = (slideWidth - shape.width) / 2;
                });
                
                await context.sync();
            }
        });
        showNotification("Objects aligned to center");
        return;
    } catch (error) {
        console.error("Error aligning center:", error);
        showNotification("Error: Could not align objects to center", "error");
        return error.code;
    }
}

async function alignLeft() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            
            if (selectedShapes.length > 0) {
                const leftPosition = Math.min(...selectedShapes.map(shape => shape.left));
                
                selectedShapes.forEach(shape => {
                    shape.left = leftPosition;
                });
                
                await context.sync();
            }
        });
        showNotification("Objects aligned to left");
        return;
    } catch (error) {
        console.error("Error aligning left:", error);
        showNotification("Error: Could not align objects to left", "error");
        return error.code;
    }
}

async function alignRight() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            
            if (selectedShapes.length > 0) {
                const rightPosition = Math.max(...selectedShapes.map(shape => shape.left + shape.width));
                
                selectedShapes.forEach(shape => {
                    shape.left = rightPosition - shape.width;
                });
                
                await context.sync();
            }
        });
        showNotification("Objects aligned to right");
        return;
    } catch (error) {
        console.error("Error aligning right:", error);
        showNotification("Error: Could not align objects to right", "error");
        return error.code;
    }
}

async function alignMiddle() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            
            if (selectedShapes.length > 0) {
                const slideHeight = 540; // Standard slide height in points
                
                selectedShapes.forEach(shape => {
                    shape.top = (slideHeight - shape.height) / 2;
                });
                
                await context.sync();
            }
        });
        showNotification("Objects aligned to middle");
        return;
    } catch (error) {
        console.error("Error aligning middle:", error);
        showNotification("Error: Could not align objects to middle", "error");
        return error.code;
    }
}

async function distributeHorizontally() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            
            if (selectedShapes.length > 2) {
                // Sort shapes by left position
                selectedShapes.sort((a, b) => a.left - b.left);
                
                const leftmost = selectedShapes[0].left;
                const rightmost = selectedShapes[selectedShapes.length - 1].left + selectedShapes[selectedShapes.length - 1].width;
                const totalSpace = rightmost - leftmost;
                const spacing = totalSpace / (selectedShapes.length - 1);
                
                for (let i = 1; i < selectedShapes.length - 1; i++) {
                    selectedShapes[i].left = leftmost + (spacing * i);
                }
                
                await context.sync();
            }
        });
        showNotification("Objects distributed horizontally");
        return;
    } catch (error) {
        console.error("Error distributing horizontally:", error);
        showNotification("Error: Could not distribute objects horizontally", "error");
        return error.code;
    }
}

async function distributeVertically() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            
            if (selectedShapes.length > 2) {
                // Sort shapes by top position
                selectedShapes.sort((a, b) => a.top - b.top);
                
                const topmost = selectedShapes[0].top;
                const bottommost = selectedShapes[selectedShapes.length - 1].top + selectedShapes[selectedShapes.length - 1].height;
                const totalSpace = bottommost - topmost;
                const spacing = totalSpace / (selectedShapes.length - 1);
                
                for (let i = 1; i < selectedShapes.length - 1; i++) {
                    selectedShapes[i].top = topmost + (spacing * i);
                }
                
                await context.sync();
            }
        });
        showNotification("Objects distributed vertically");
        return;
    } catch (error) {
        console.error("Error distributing vertically:", error);
        showNotification("Error: Could not distribute objects vertically", "error");
        return error.code;
    }
}

// Professional Element Functions
async function insertFootnote() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            
            // Create footnote textbox at bottom of slide
            const footnote = slide.shapes.addTextBox(
                "Source: Professional analysis",
                {
                    left: 50,
                    top: 480, // Near bottom of slide
                    width: 620,
                    height: 30
                }
            );
            
            // Style the footnote
            footnote.textFrame.textRange.font.name = "Calibri";
            footnote.textFrame.textRange.font.size = 10;
            footnote.textFrame.textRange.font.color = "#666666";
            footnote.textFrame.textRange.font.italic = true;
            
            await context.sync();
        });
        showNotification("Professional footnote inserted");
        return;
    } catch (error) {
        console.error("Error inserting footnote:", error);
        showNotification("Error: Could not insert footnote", "error");
        return error.code;
    }
}

async function insertLegend() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            
            // Create legend box
            const legend = slide.shapes.addGeometricShape(
                PowerPoint.GeometricShapeType.rectangle,
                {
                    left: 500,
                    top: 100,
                    width: 150,
                    height: 100
                }
            );
            
            // Style the legend
            legend.fill.setSolidColor("#F5F5F5");
            legend.line.color = "#CCCCCC";
            legend.line.weight = 1;
            
            legend.textFrame.textRange.text = "Legend\n• Item 1\n• Item 2\n• Item 3";
            legend.textFrame.textRange.font.name = "Calibri";
            legend.textFrame.textRange.font.size = 11;
            legend.textFrame.textRange.font.color = "#333333";
            
            await context.sync();
        });
        showNotification("Professional legend inserted");
        return;
    } catch (error) {
        console.error("Error inserting legend:", error);
        showNotification("Error: Could not insert legend", "error");
        return error.code;
    }
}

async function insertSticker() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            
            // Create circular sticker
            const sticker = slide.shapes.addGeometricShape(
                PowerPoint.GeometricShapeType.ellipse,
                {
                    left: 200,
                    top: 200,
                    width: 80,
                    height: 80
                }
            );
            
            // Style the sticker
            sticker.fill.setSolidColor("#00A651"); // Professional green
            sticker.line.color = "#FFFFFF";
            sticker.line.weight = 2;
            
            sticker.textFrame.textRange.text = "1";
            sticker.textFrame.textRange.font.name = "Calibri";
            sticker.textFrame.textRange.font.size = 24;
            sticker.textFrame.textRange.font.color = "#FFFFFF";
            sticker.textFrame.textRange.font.bold = true;
            sticker.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
            
            await context.sync();
        });
        showNotification("Professional sticker inserted");
        return;
    } catch (error) {
        console.error("Error inserting sticker:", error);
        showNotification("Error: Could not insert sticker", "error");
        return error.code;
    }
}

async function cycleAccentColors() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            const professionalColors = ["#00A651", "#0073E6", "#FF6B35", "#8E44AD", "#F39C12"];
            let colorIndex = 0;
            
            selectedShapes.forEach(shape => {
                if (shape.fill) {
                    shape.fill.setSolidColor(professionalColors[colorIndex % professionalColors.length]);
                    colorIndex++;
                }
            });
            
            await context.sync();
        });
        showNotification("Accent colors cycled");
        return;
    } catch (error) {
        console.error("Error cycling accent colors:", error);
        showNotification("Error: Could not cycle accent colors", "error");
        return error.code;
    }
}

// Utility Functions
async function greenPrint() {
    try {
        showNotification("Green Print feature would export presentation with Professional green theme");
        // In a real implementation, this would:
        // 1. Apply green theme to all slides
        // 2. Export as PDF
        // 3. Restore original theme
        return;
    } catch (error) {
        console.error("Error with green print:", error);
        showNotification("Error: Could not execute green print", "error");
        return error.code;
    }
}

async function showQuickKeys() {
    try {
        // Show a dialog with all available shortcuts
        const shortcuts = `
Professional PowerPoint Shortcuts:

Text & Objects:
• Ctrl+Alt+T - Paste unformatted text
• Shift+Alt+Z - Text to autoshape
• Alt+Ctrl+J - Split/join textboxes

Alignment:
• Shift+Alt+E - Make same width
• Ctrl+Alt+C - Align center
• Ctrl+Alt+L - Align left
• Ctrl+Alt+R - Align right
• Ctrl+Alt+M - Align middle
• Alt+Shift+H - Distribute horizontally
• Alt+Shift+V - Distribute vertically

Professional Elements:
• Ctrl+Alt+F - Insert footnote
• Ctrl+Alt+G - Insert legend
• Ctrl+Alt+S - Insert sticker
• Shift+Alt+A - Cycle accent colors

Utilities:
• Ctrl+Alt+P - Green print
• Ctrl+Alt+Q - Show this help
• Ctrl+Alt+Y - Reset elements
• Ctrl+Alt+K - Show shortcuts panel
        `;
        
        showNotification(shortcuts, "info", 10000);
        return;
    } catch (error) {
        console.error("Error showing quick keys:", error);
        return error.code;
    }
}

async function resetElements() {
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            
            selectedShapes.forEach(shape => {
                // Reset to default formatting
                if (shape.textFrame && shape.textFrame.hasText) {
                    shape.textFrame.textRange.font.name = "Calibri";
                    shape.textFrame.textRange.font.size = 18;
                    shape.textFrame.textRange.font.bold = false;
                    shape.textFrame.textRange.font.italic = false;
                    shape.textFrame.textRange.font.color = "#000000";
                }
                
                if (shape.fill) {
                    shape.fill.clear();
                }
                
                if (shape.line) {
                    shape.line.color = "#000000";
                    shape.line.weight = 1;
                }
            });
            
            await context.sync();
        });
        showNotification("Elements reset to default formatting");
        return;
    } catch (error) {
        console.error("Error resetting elements:", error);
        showNotification("Error: Could not reset elements", "error");
        return error.code;
    }
}

async function showTaskpane() {
    try {
        await Office.addin.showAsTaskpane();
        return;
    } catch (error) {
        console.error("Error showing taskpane:", error);
        return error.code;
    }
}

// Utility function to show notifications
function showNotification(message, type = "info", duration = 3000) {
    console.log(`[${type.toUpperCase()}] ${message}`);
    
    // In a real implementation, this would show a toast notification
    // For now, we'll use console logging and could implement a simple overlay
    
    if (typeof Office !== 'undefined' && Office.context && Office.context.ui) {
        // Use Office UI notification if available
        Office.context.ui.displayDialogAsync(
            `data:text/html,<html><body><h3>${type.toUpperCase()}</h3><p>${message}</p></body></html>`,
            { height: 30, width: 50 }
        );
    }
}

