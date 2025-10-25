/*
 * PowerPoint Shortcuts - Task Pane JavaScript
 * User interface interactions and functionality
 */

// Initialize Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("PowerPoint Shortcuts Task Pane loaded");
        initializeTaskPane();
    }
});

// Initialize task pane functionality
function initializeTaskPane() {
    updateStatus("PowerPoint Shortcuts ready");
    
    // Set up event listeners
    setupEventListeners();
    
    // Show default tab
    showTab('shortcuts');
}

// Set up event listeners
function setupEventListeners() {
    // Tab switching
    document.querySelectorAll('.tab-button').forEach(button => {
        button.addEventListener('click', (e) => {
            const tabName = e.target.getAttribute('onclick').match(/'([^']+)'/)[1];
            showTab(tabName);
        });
    });
    
    // Keyboard shortcuts for quick access
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.altKey) {
            switch (e.key.toLowerCase()) {
                case 'h':
                    e.preventDefault();
                    showTab('help');
                    break;
                case 's':
                    e.preventDefault();
                    showTab('shortcuts');
                    break;
                case 'e':
                    e.preventDefault();
                    showTab('elements');
                    break;
                case 'f':
                    e.preventDefault();
                    showTab('formatting');
                    break;
            }
        }
    });
}

// Tab switching functionality
function showTab(tabName) {
    // Hide all tab contents
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // Remove active class from all tab buttons
    document.querySelectorAll('.tab-button').forEach(button => {
        button.classList.remove('active');
    });
    
    // Show selected tab content
    const selectedTab = document.getElementById(`${tabName}-tab`);
    if (selectedTab) {
        selectedTab.classList.add('active');
    }
    
    // Add active class to selected tab button
    const selectedButton = document.querySelector(`[onclick="showTab('${tabName}')"]`);
    if (selectedButton) {
        selectedButton.classList.add('active');
    }
    
    updateStatus(`Viewing ${tabName} tab`);
}

// Execute shortcut functions
async function executeShortcut(actionId) {
    updateStatus(`Executing ${actionId}...`);
    
    try {
        // Map action IDs to their corresponding functions
        const actionMap = {
            'PasteUnformatted': pasteUnformattedText,
            'TextToAutoshape': textToAutoshape,
            'SplitJoinTextboxes': splitJoinTextboxes,
            'MakeSameWidth': makeSameWidth,
            'AlignCenter': alignCenter,
            'AlignLeft': alignLeft,
            'AlignRight': alignRight,
            'AlignMiddle': alignMiddle,
            'DistributeHorizontally': distributeHorizontally,
            'DistributeVertically': distributeVertically,
            'InsertFootnote': insertFootnote,
            'InsertLegend': insertLegend,
            'InsertSticker': insertSticker,
            'CycleAccentColors': cycleAccentColors,
            'GreenPrint': greenPrint,
            'ShowQuickKeys': showQuickKeys,
            'ResetElements': resetElements
        };
        
        const actionFunction = actionMap[actionId];
        if (actionFunction) {
            await actionFunction();
            updateStatus(`${actionId} completed successfully`);
            showNotification(`${actionId} executed successfully`);
        } else {
            throw new Error(`Unknown action: ${actionId}`);
        }
    } catch (error) {
        console.error(`Error executing ${actionId}:`, error);
        updateStatus(`Error executing ${actionId}`);
        showNotification(`Error: Could not execute ${actionId}`, 'error');
    }
}

// Professional element creation functions
async function createSlideTemplate(templateType) {
    updateStatus(`Creating ${templateType} template...`);
    
    try {
        await ProfessionalFunctions.createSlideTemplate(templateType);
        updateStatus(`${templateType} template created`);
        showNotification(`${templateType} template created successfully`);
    } catch (error) {
        console.error(`Error creating ${templateType} template:`, error);
        updateStatus(`Error creating ${templateType} template`);
        showNotification(`Error: Could not create ${templateType} template`, 'error');
    }
}

async function createChartPlaceholder() {
    updateStatus('Creating chart placeholder...');
    
    try {
        await ProfessionalFunctions.createChartPlaceholder();
        updateStatus('Chart placeholder created');
        showNotification('Chart placeholder created successfully');
    } catch (error) {
        console.error('Error creating chart placeholder:', error);
        updateStatus('Error creating chart placeholder');
        showNotification('Error: Could not create chart placeholder', 'error');
    }
}

// Color application function
async function applyColor(color) {
    updateStatus(`Applying color ${color}...`);
    
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            // Apply color to all shapes (in real implementation, would apply to selected shapes)
            const selectedShapes = shapes.items.filter(shape => shape.id);
            
            selectedShapes.forEach(shape => {
                if (shape.fill) {
                    shape.fill.setSolidColor(color);
                }
            });
            
            await context.sync();
        });
        
        updateStatus(`Color ${color} applied`);
        showNotification(`Color applied successfully`);
    } catch (error) {
        console.error('Error applying color:', error);
        updateStatus('Error applying color');
        showNotification('Error: Could not apply color', 'error');
    }
}

// Professional grid application
async function applyProfessionalGrid() {
    updateStatus('Applying Professional grid...');
    
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            const selectedShapes = shapes.items.filter(shape => shape.id);
            await ProfessionalFunctions.applyProfessionalGrid(selectedShapes);
        });
        
        updateStatus('Professional grid applied');
        showNotification('Professional grid applied successfully');
    } catch (error) {
        console.error('Error applying Professional grid:', error);
        updateStatus('Error applying Professional grid');
        showNotification('Error: Could not apply Professional grid', 'error');
    }
}

// Shortcut function implementations (these would normally be in commands.js)
async function pasteUnformattedText() {
    return await PowerPoint.run(async (context) => {
        // Implementation would go here
        // For demo purposes, just show a message
        showNotification("Paste unformatted text functionality");
    });
}

async function textToAutoshape() {
    return await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        
        // Create a sample autoshape
        const shape = slide.shapes.addGeometricShape(
            PowerPoint.GeometricShapeType.roundRectangle,
            {
                left: 200,
                top: 200,
                width: 150,
                height: 50
            }
        );
        
        shape.fill.setSolidColor(ProfessionalFunctions.COLORS.primary);
        shape.textFrame.textRange.text = "Sample Text";
        shape.textFrame.textRange.font.color = ProfessionalFunctions.COLORS.white;
        shape.textFrame.textRange.font.bold = true;
        
        await context.sync();
    });
}

async function splitJoinTextboxes() {
    showNotification("Split/join textboxes functionality");
}

async function makeSameWidth() {
    return await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        shapes.load("items");
        
        await context.sync();
        
        const selectedShapes = shapes.items.filter(shape => shape.id);
        
        if (selectedShapes.length > 1) {
            const referenceWidth = selectedShapes[0].width;
            
            for (let i = 1; i < selectedShapes.length; i++) {
                selectedShapes[i].width = referenceWidth;
            }
            
            await context.sync();
        }
    });
}

async function alignCenter() {
    return await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        shapes.load("items");
        
        await context.sync();
        
        const selectedShapes = shapes.items.filter(shape => shape.id);
        const slideWidth = 720;
        
        selectedShapes.forEach(shape => {
            shape.left = (slideWidth - shape.width) / 2;
        });
        
        await context.sync();
    });
}

async function alignLeft() {
    return await PowerPoint.run(async (context) => {
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
}

async function alignRight() {
    return await PowerPoint.run(async (context) => {
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
}

async function alignMiddle() {
    return await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        shapes.load("items");
        
        await context.sync();
        
        const selectedShapes = shapes.items.filter(shape => shape.id);
        const slideHeight = 540;
        
        selectedShapes.forEach(shape => {
            shape.top = (slideHeight - shape.height) / 2;
        });
        
        await context.sync();
    });
}

async function distributeHorizontally() {
    showNotification("Distribute horizontally functionality");
}

async function distributeVertically() {
    showNotification("Distribute vertically functionality");
}

async function insertFootnote() {
    return await ProfessionalFunctions.createFootnote();
}

async function insertLegend() {
    return await ProfessionalFunctions.createLegend();
}

async function insertSticker() {
    const nextNumber = await ProfessionalFunctions.getNextStickerNumber();
    return await ProfessionalFunctions.createSticker(nextNumber);
}

async function cycleAccentColors() {
    return await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        shapes.load("items");
        
        await context.sync();
        
        const selectedShapes = shapes.items.filter(shape => shape.id);
        await ProfessionalFunctions.applyColorPalette(selectedShapes);
        
        await context.sync();
    });
}

async function greenPrint() {
    return await ProfessionalFunctions.exportWithGreenTheme();
}

async function showQuickKeys() {
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

Task Pane Navigation:
• Ctrl+Alt+H - Help tab
• Ctrl+Alt+S - Shortcuts tab
• Ctrl+Alt+E - Elements tab
• Ctrl+Alt+F - Formatting tab
    `;
    
    showNotification(shortcuts, 'info', 10000);
}

async function resetElements() {
    return await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        shapes.load("items");
        
        await context.sync();
        
        const selectedShapes = shapes.items.filter(shape => shape.id);
        await ProfessionalFunctions.resetToProfessionalDefaults(selectedShapes);
        
        await context.sync();
    });
}

// Utility functions
function updateStatus(message) {
    const statusElement = document.getElementById('status-message');
    if (statusElement) {
        statusElement.textContent = message;
    }
    console.log(`Status: ${message}`);
}

function showNotification(message, type = 'info', duration = 3000) {
    // Remove existing notifications
    const existingNotifications = document.querySelectorAll('.notification');
    existingNotifications.forEach(notification => {
        notification.remove();
    });
    
    // Create new notification
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    
    // Add to document
    document.body.appendChild(notification);
    
    // Auto-remove after duration
    setTimeout(() => {
        if (notification.parentNode) {
            notification.parentNode.removeChild(notification);
        }
    }, duration);
    
    // Also log to console
    console.log(`[${type.toUpperCase()}] ${message}`);
}

// Error handling
window.addEventListener('error', (event) => {
    console.error('Task pane error:', event.error);
    showNotification('An unexpected error occurred', 'error');
    updateStatus('Error occurred');
});

// Unhandled promise rejection handling
window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled promise rejection:', event.reason);
    showNotification('An unexpected error occurred', 'error');
    updateStatus('Error occurred');
});

// Export functions for global access
window.showTab = showTab;
window.executeShortcut = executeShortcut;
window.createSlideTemplate = createSlideTemplate;
window.createChartPlaceholder = createChartPlaceholder;
window.applyColor = applyColor;
window.applyProfessionalGrid = applyProfessionalGrid;

