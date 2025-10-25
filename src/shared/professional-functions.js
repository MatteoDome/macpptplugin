/*
 * Professional Functions
 * Business logic for professional elements and styling
 */

class ProfessionalFunctions {
    
    // Professional brand colors
    static COLORS = {
        primary: "#0078D4",      // Blue
        secondary: "#107C10",    // Green
        accent1: "#FF4B4B",      // Red
        accent2: "#8E44AD",      // Purple
        accent3: "#F39C12",      // Orange
        neutral1: "#2C3E50",     // Dark Gray
        neutral2: "#7F8C8D",     // Medium Gray
        neutral3: "#BDC3C7",     // Light Gray
        neutral4: "#ECF0F1",     // Very Light Gray
        white: "#FFFFFF",
        black: "#000000"
    };
    
    // Professional typography settings
    static TYPOGRAPHY = {
        primaryFont: "Calibri",
        secondaryFont: "Arial",
        titleSize: 24,
        subtitleSize: 18,
        bodySize: 14,
        footnoteSize: 10,
        captionSize: 12
    };
    
    // Standard Professional slide dimensions and margins
    static LAYOUT = {
        slideWidth: 720,
        slideHeight: 540,
        marginTop: 50,
        marginBottom: 50,
        marginLeft: 50,
        marginRight: 50,
        contentWidth: 620,
        contentHeight: 440
    };
    
    // Create a Professional-styled footnote
    static async createFootnote(text = "Source: Professional analysis", position = "bottom-left") {
        const options = {
            left: ProfessionalFunctions.LAYOUT.marginLeft,
            top: ProfessionalFunctions.LAYOUT.slideHeight - ProfessionalFunctions.LAYOUT.marginBottom,
            width: ProfessionalFunctions.LAYOUT.contentWidth,
            height: 25,
            fontName: ProfessionalFunctions.TYPOGRAPHY.primaryFont,
            fontSize: ProfessionalFunctions.TYPOGRAPHY.footnoteSize,
            fontColor: ProfessionalFunctions.COLORS.neutral2
        };
        
        return await PowerPointUtils.createTextBox(text, options);
    }
    
    // Create a Professional-styled legend box
    static async createLegend(items = ["Item 1", "Item 2", "Item 3"], position = "top-right") {
        const legendText = "Legend\n" + items.map(item => `• ${item}`).join("\n");
        
        const options = {
            left: ProfessionalFunctions.LAYOUT.slideWidth - 200 - ProfessionalFunctions.LAYOUT.marginRight,
            top: ProfessionalFunctions.LAYOUT.marginTop,
            width: 180,
            height: 80 + (items.length * 15),
            fontName: ProfessionalFunctions.TYPOGRAPHY.primaryFont,
            fontSize: ProfessionalFunctions.TYPOGRAPHY.captionSize,
            fontColor: ProfessionalFunctions.COLORS.neutral1,
            backgroundColor: ProfessionalFunctions.COLORS.neutral4
        };
        
        const legend = await PowerPointUtils.createTextBox(legendText, options);
        
        // Add border
        return await PowerPoint.run(async (context) => {
            legend.line.color = ProfessionalFunctions.COLORS.neutral3;
            legend.line.weight = 1;
            await context.sync();
            return legend;
        });
    }
    
    // Create a Professional-styled sticker (circular numbered element)
    static async createSticker(number = "1", color = null) {
        const stickerColor = color || ProfessionalFunctions.COLORS.primary;
        
        const options = {
            left: 200,
            top: 200,
            width: 60,
            height: 60,
            fillColor: stickerColor,
            lineColor: ProfessionalFunctions.COLORS.white,
            lineWeight: 2
        };
        
        const sticker = await PowerPointUtils.createShape(
            PowerPoint.GeometricShapeType.ellipse,
            options
        );
        
        return await PowerPoint.run(async (context) => {
            sticker.textFrame.textRange.text = number.toString();
            sticker.textFrame.textRange.font.name = ProfessionalFunctions.TYPOGRAPHY.primaryFont;
            sticker.textFrame.textRange.font.size = 18;
            sticker.textFrame.textRange.font.color = ProfessionalFunctions.COLORS.white;
            sticker.textFrame.textRange.font.bold = true;
            sticker.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
            sticker.textFrame.verticalAlignment = PowerPoint.VerticalAlignment.middle;
            
            await context.sync();
            return sticker;
        });
    }
    
    // Apply Professional color palette to selected shapes
    static async applyColorPalette(shapes, startColorIndex = 0) {
        const colors = [
            ProfessionalFunctions.COLORS.primary,
            ProfessionalFunctions.COLORS.secondary,
            ProfessionalFunctions.COLORS.accent1,
            ProfessionalFunctions.COLORS.accent2,
            ProfessionalFunctions.COLORS.accent3
        ];
        
        return await PowerPoint.run(async (context) => {
            shapes.forEach((shape, index) => {
                const colorIndex = (startColorIndex + index) % colors.length;
                if (shape.fill) {
                    shape.fill.setSolidColor(colors[colorIndex]);
                }
            });
            
            await context.sync();
        });
    }
    
    // Create a Professional-styled text box with proper formatting
    static async createProfessionalTextBox(text, style = "body") {
        let options = {
            fontName: ProfessionalFunctions.TYPOGRAPHY.primaryFont,
            fontColor: ProfessionalFunctions.COLORS.neutral1
        };
        
        switch (style) {
            case "title":
                options.fontSize = ProfessionalFunctions.TYPOGRAPHY.titleSize;
                options.fontColor = ProfessionalFunctions.COLORS.primary;
                break;
            case "subtitle":
                options.fontSize = ProfessionalFunctions.TYPOGRAPHY.subtitleSize;
                options.fontColor = ProfessionalFunctions.COLORS.neutral1;
                break;
            case "body":
                options.fontSize = ProfessionalFunctions.TYPOGRAPHY.bodySize;
                break;
            case "caption":
                options.fontSize = ProfessionalFunctions.TYPOGRAPHY.captionSize;
                options.fontColor = ProfessionalFunctions.COLORS.neutral2;
                break;
        }
        
        return await PowerPointUtils.createTextBox(text, options);
    }
    
    // Convert text to a Professional-styled autoshape
    static async convertTextToProfessionalShape(text, shapeType = "roundedRectangle") {
        const shapeTypeMap = {
            "roundedRectangle": PowerPoint.GeometricShapeType.roundRectangle,
            "rectangle": PowerPoint.GeometricShapeType.rectangle,
            "ellipse": PowerPoint.GeometricShapeType.ellipse,
            "hexagon": PowerPoint.GeometricShapeType.hexagon
        };
        
        const options = {
            left: 200,
            top: 200,
            width: Math.max(text.length * 8, 120), // Dynamic width based on text length
            height: 40,
            fillColor: ProfessionalFunctions.COLORS.primary,
            lineColor: ProfessionalFunctions.COLORS.white,
            lineWeight: 1
        };
        
        const shape = await PowerPointUtils.createShape(
            shapeTypeMap[shapeType] || PowerPoint.GeometricShapeType.roundRectangle,
            options
        );
        
        return await PowerPoint.run(async (context) => {
            shape.textFrame.textRange.text = text;
            shape.textFrame.textRange.font.name = ProfessionalFunctions.TYPOGRAPHY.primaryFont;
            shape.textFrame.textRange.font.size = ProfessionalFunctions.TYPOGRAPHY.bodySize;
            shape.textFrame.textRange.font.color = ProfessionalFunctions.COLORS.white;
            shape.textFrame.textRange.font.bold = true;
            shape.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
            shape.textFrame.verticalAlignment = PowerPoint.VerticalAlignment.middle;
            
            await context.sync();
            return shape;
        });
    }
    
    // Apply Professional grid alignment to shapes
    static async applyProfessionalGrid(shapes) {
        const gridSize = 20; // 20-point grid
        
        return await PowerPoint.run(async (context) => {
            shapes.forEach(shape => {
                // Snap to grid
                shape.left = Math.round(shape.left / gridSize) * gridSize;
                shape.top = Math.round(shape.top / gridSize) * gridSize;
                shape.width = Math.round(shape.width / gridSize) * gridSize;
                shape.height = Math.round(shape.height / gridSize) * gridSize;
            });
            
            await context.sync();
        });
    }
    
    // Reset shapes to Professional default formatting
    static async resetToProfessionalDefaults(shapes) {
        return await PowerPoint.run(async (context) => {
            shapes.forEach(shape => {
                // Reset text formatting
                if (shape.textFrame && shape.textFrame.hasText) {
                    shape.textFrame.textRange.font.name = ProfessionalFunctions.TYPOGRAPHY.primaryFont;
                    shape.textFrame.textRange.font.size = ProfessionalFunctions.TYPOGRAPHY.bodySize;
                    shape.textFrame.textRange.font.color = ProfessionalFunctions.COLORS.neutral1;
                    shape.textFrame.textRange.font.bold = false;
                    shape.textFrame.textRange.font.italic = false;
                    shape.textFrame.textRange.font.underline = PowerPoint.UnderlineType.none;
                }
                
                // Reset shape formatting
                if (shape.fill) {
                    shape.fill.setSolidColor(ProfessionalFunctions.COLORS.white);
                }
                
                if (shape.line) {
                    shape.line.color = ProfessionalFunctions.COLORS.neutral3;
                    shape.line.weight = 1;
                }
            });
            
            await context.sync();
        });
    }
    
    // Create a Professional-styled chart placeholder
    static async createChartPlaceholder(chartType = "column") {
        const options = {
            left: ProfessionalFunctions.LAYOUT.marginLeft,
            top: ProfessionalFunctions.LAYOUT.marginTop + 60,
            width: ProfessionalFunctions.LAYOUT.contentWidth * 0.7,
            height: ProfessionalFunctions.LAYOUT.contentHeight * 0.6,
            fillColor: ProfessionalFunctions.COLORS.neutral4,
            lineColor: ProfessionalFunctions.COLORS.neutral3,
            lineWeight: 1
        };
        
        const placeholder = await PowerPointUtils.createShape(
            PowerPoint.GeometricShapeType.rectangle,
            options
        );
        
        return await PowerPoint.run(async (context) => {
            placeholder.textFrame.textRange.text = `${chartType.toUpperCase()} CHART\nPlaceholder`;
            placeholder.textFrame.textRange.font.name = ProfessionalFunctions.TYPOGRAPHY.primaryFont;
            placeholder.textFrame.textRange.font.size = ProfessionalFunctions.TYPOGRAPHY.subtitleSize;
            placeholder.textFrame.textRange.font.color = ProfessionalFunctions.COLORS.neutral2;
            placeholder.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
            placeholder.textFrame.verticalAlignment = PowerPoint.VerticalAlignment.middle;
            
            await context.sync();
            return placeholder;
        });
    }
    
    // Generate a Professional-styled slide template
    static async createSlideTemplate(templateType = "content") {
        return await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            
            // Clear existing content
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            
            // Add title placeholder
            const title = await ProfessionalFunctions.createProfessionalTextBox(
                "Slide Title",
                "title"
            );
            title.left = ProfessionalFunctions.LAYOUT.marginLeft;
            title.top = ProfessionalFunctions.LAYOUT.marginTop;
            title.width = ProfessionalFunctions.LAYOUT.contentWidth;
            title.height = 40;
            
            if (templateType === "content") {
                // Add content area
                const content = await ProfessionalFunctions.createProfessionalTextBox(
                    "• Bullet point 1\n• Bullet point 2\n• Bullet point 3",
                    "body"
                );
                content.left = ProfessionalFunctions.LAYOUT.marginLeft;
                content.top = ProfessionalFunctions.LAYOUT.marginTop + 80;
                content.width = ProfessionalFunctions.LAYOUT.contentWidth * 0.6;
                content.height = 200;
            } else if (templateType === "chart") {
                // Add chart placeholder
                await ProfessionalFunctions.createChartPlaceholder();
            }
            
            // Add footnote
            await ProfessionalFunctions.createFootnote();
            
            await context.sync();
        });
    }
    
    // Export current presentation with Professional green theme
    static async exportWithGreenTheme() {
        // This would be implemented to:
        // 1. Save current theme
        // 2. Apply green theme to all slides
        // 3. Export as PDF
        // 4. Restore original theme
        
        PowerPointUtils.showNotification(
            "Green Print: This feature would export the presentation with Professional green theme applied",
            "info",
            5000
        );
        
        // Placeholder implementation
        return Promise.resolve();
    }
    
    // Get next available sticker number
    static async getNextStickerNumber() {
        return await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            // Find existing stickers and determine next number
            let maxNumber = 0;
            shapes.items.forEach(shape => {
                if (shape.type === PowerPoint.ShapeType.geometricShape && 
                    shape.geometricShapeType === PowerPoint.GeometricShapeType.ellipse &&
                    shape.textFrame && shape.textFrame.hasText) {
                    const text = shape.textFrame.textRange.text.trim();
                    const number = parseInt(text);
                    if (!isNaN(number) && number > maxNumber) {
                        maxNumber = number;
                    }
                }
            });
            
            return (maxNumber + 1).toString();
        });
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ProfessionalFunctions;
}

