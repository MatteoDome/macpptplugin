/*
 * PowerPoint Utilities
 * Common functions for working with PowerPoint objects
 */

class PowerPointUtils {
    
    // Get currently selected shapes
    static async getSelectedShapes() {
        return await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            // In a real implementation, we would need to track which shapes are selected
            // For now, return all shapes as a placeholder
            return shapes.items;
        });
    }
    
    // Get slide dimensions
    static async getSlideDimensions() {
        return await PowerPoint.run(async (context) => {
            const presentation = context.presentation;
            presentation.load("slideMaster");
            
            await context.sync();
            
            // Standard PowerPoint slide dimensions in points
            return {
                width: 720,  // 10 inches at 72 DPI
                height: 540  // 7.5 inches at 72 DPI
            };
        });
    }
    
    // Create a text box with specified properties
    static async createTextBox(text, options = {}) {
        return await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            
            const defaultOptions = {
                left: 100,
                top: 100,
                width: 200,
                height: 50,
                fontName: "Calibri",
                fontSize: 18,
                fontColor: "#000000",
                backgroundColor: "transparent"
            };
            
            const config = { ...defaultOptions, ...options };
            
            const textBox = slide.shapes.addTextBox(
                text,
                {
                    left: config.left,
                    top: config.top,
                    width: config.width,
                    height: config.height
                }
            );
            
            // Apply formatting
            textBox.textFrame.textRange.font.name = config.fontName;
            textBox.textFrame.textRange.font.size = config.fontSize;
            textBox.textFrame.textRange.font.color = config.fontColor;
            
            if (config.backgroundColor !== "transparent") {
                textBox.fill.setSolidColor(config.backgroundColor);
            }
            
            await context.sync();
            return textBox;
        });
    }
    
    // Create a shape with specified properties
    static async createShape(shapeType, options = {}) {
        return await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            
            const defaultOptions = {
                left: 100,
                top: 100,
                width: 100,
                height: 100,
                fillColor: "#FFFFFF",
                lineColor: "#000000",
                lineWeight: 1
            };
            
            const config = { ...defaultOptions, ...options };
            
            const shape = slide.shapes.addGeometricShape(
                shapeType,
                {
                    left: config.left,
                    top: config.top,
                    width: config.width,
                    height: config.height
                }
            );
            
            // Apply formatting
            shape.fill.setSolidColor(config.fillColor);
            shape.line.color = config.lineColor;
            shape.line.weight = config.lineWeight;
            
            await context.sync();
            return shape;
        });
    }
    
    // Apply consistent formatting to multiple shapes
    static async applyConsistentFormatting(shapes, formatting) {
        return await PowerPoint.run(async (context) => {
            shapes.forEach(shape => {
                if (formatting.fillColor && shape.fill) {
                    shape.fill.setSolidColor(formatting.fillColor);
                }
                
                if (formatting.lineColor && shape.line) {
                    shape.line.color = formatting.lineColor;
                }
                
                if (formatting.lineWeight && shape.line) {
                    shape.line.weight = formatting.lineWeight;
                }
                
                if (formatting.fontName && shape.textFrame && shape.textFrame.hasText) {
                    shape.textFrame.textRange.font.name = formatting.fontName;
                }
                
                if (formatting.fontSize && shape.textFrame && shape.textFrame.hasText) {
                    shape.textFrame.textRange.font.size = formatting.fontSize;
                }
                
                if (formatting.fontColor && shape.textFrame && shape.textFrame.hasText) {
                    shape.textFrame.textRange.font.color = formatting.fontColor;
                }
            });
            
            await context.sync();
        });
    }
    
    // Calculate bounding box for multiple shapes
    static calculateBoundingBox(shapes) {
        if (shapes.length === 0) return null;
        
        let minLeft = Infinity;
        let minTop = Infinity;
        let maxRight = -Infinity;
        let maxBottom = -Infinity;
        
        shapes.forEach(shape => {
            minLeft = Math.min(minLeft, shape.left);
            minTop = Math.min(minTop, shape.top);
            maxRight = Math.max(maxRight, shape.left + shape.width);
            maxBottom = Math.max(maxBottom, shape.top + shape.height);
        });
        
        return {
            left: minLeft,
            top: minTop,
            width: maxRight - minLeft,
            height: maxBottom - minTop,
            right: maxRight,
            bottom: maxBottom
        };
    }
    
    // Align shapes relative to each other
    static async alignShapes(shapes, alignment) {
        return await PowerPoint.run(async (context) => {
            if (shapes.length < 2) return;
            
            const boundingBox = PowerPointUtils.calculateBoundingBox(shapes);
            
            switch (alignment) {
                case 'left':
                    shapes.forEach(shape => {
                        shape.left = boundingBox.left;
                    });
                    break;
                    
                case 'right':
                    shapes.forEach(shape => {
                        shape.left = boundingBox.right - shape.width;
                    });
                    break;
                    
                case 'center':
                    const centerX = boundingBox.left + boundingBox.width / 2;
                    shapes.forEach(shape => {
                        shape.left = centerX - shape.width / 2;
                    });
                    break;
                    
                case 'top':
                    shapes.forEach(shape => {
                        shape.top = boundingBox.top;
                    });
                    break;
                    
                case 'bottom':
                    shapes.forEach(shape => {
                        shape.top = boundingBox.bottom - shape.height;
                    });
                    break;
                    
                case 'middle':
                    const centerY = boundingBox.top + boundingBox.height / 2;
                    shapes.forEach(shape => {
                        shape.top = centerY - shape.height / 2;
                    });
                    break;
            }
            
            await context.sync();
        });
    }
    
    // Distribute shapes evenly
    static async distributeShapes(shapes, direction) {
        return await PowerPoint.run(async (context) => {
            if (shapes.length < 3) return;
            
            if (direction === 'horizontal') {
                // Sort by left position
                shapes.sort((a, b) => a.left - b.left);
                
                const leftmost = shapes[0].left;
                const rightmost = shapes[shapes.length - 1].left + shapes[shapes.length - 1].width;
                const totalSpace = rightmost - leftmost;
                const spacing = totalSpace / (shapes.length - 1);
                
                for (let i = 1; i < shapes.length - 1; i++) {
                    shapes[i].left = leftmost + (spacing * i);
                }
            } else if (direction === 'vertical') {
                // Sort by top position
                shapes.sort((a, b) => a.top - b.top);
                
                const topmost = shapes[0].top;
                const bottommost = shapes[shapes.length - 1].top + shapes[shapes.length - 1].height;
                const totalSpace = bottommost - topmost;
                const spacing = totalSpace / (shapes.length - 1);
                
                for (let i = 1; i < shapes.length - 1; i++) {
                    shapes[i].top = topmost + (spacing * i);
                }
            }
            
            await context.sync();
        });
    }
    
    // Get current slide index
    static async getCurrentSlideIndex() {
        return await PowerPoint.run(async (context) => {
            const presentation = context.presentation;
            const slides = presentation.slides;
            slides.load("items");
            
            await context.sync();
            
            // In a real implementation, we would track the current slide
            // For now, return 0 as placeholder
            return 0;
        });
    }
    
    // Show a notification to the user
    static showNotification(message, type = "info", duration = 3000) {
        console.log(`[${type.toUpperCase()}] ${message}`);
        
        // Create a simple notification overlay
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: ${type === 'error' ? '#ff4444' : type === 'warning' ? '#ffaa00' : '#00aa44'};
            color: white;
            padding: 12px 20px;
            border-radius: 4px;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 14px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            z-index: 10000;
            max-width: 300px;
            word-wrap: break-word;
        `;
        notification.textContent = message;
        
        document.body.appendChild(notification);
        
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, duration);
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PowerPointUtils;
}

