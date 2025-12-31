// assets/app.js

document.getElementById('researchForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const button = e.target.querySelector('button[type="submit"]');
    const messageDiv = document.getElementById('message');
    
    // Show loading state
    button.disabled = true;
    button.classList.add('loading');
    button.textContent = 'Generating...';
    messageDiv.className = 'message';
    messageDiv.textContent = '';
    
    try {
        const title = document.getElementById('title').value;
        const content = document.getElementById('content').value;
        const imageFiles = document.getElementById('imageUpload').files;
        
        // Create paragraphs from content (split by line breaks)
        const contentParagraphs = content.split('\n').filter(line => line.trim()).map(line => 
            new docx.Paragraph({
                text: line,
                spacing: { after: 200 }
            })
        );
        
        // Process images
        const imageParagraphs = await addImages(imageFiles);
        
        // Create Word document
        const doc = new docx.Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1440,    // 1 inch = 1440 twentieths of a point
                            right: 1440,
                            bottom: 1440,
                            left: 1440,
                        }
                    }
                },
                children: [
                    // Title
                    new docx.Paragraph({
                        text: title,
                        heading: docx.HeadingLevel.HEADING_1,
                        spacing: { after: 400 },
                        alignment: docx.AlignmentType.CENTER
                    }),
                    
                    // Content
                    ...contentParagraphs,
                    
                    // Images section
                    ...(imageParagraphs.length > 0 ? [
                        new docx.Paragraph({
                            text: "Figures and Graphs",
                            heading: docx.HeadingLevel.HEADING_2,
                            spacing: { before: 400, after: 200 }
                        }),
                        ...imageParagraphs
                    ] : [])
                ]
            }]
        });
        
        // Generate and download
        const blob = await docx.Packer.toBlob(doc);
        const fileName = `${title.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_research.docx`;
        saveAs(blob, fileName);
        
        // Show success message
        messageDiv.className = 'message success';
        messageDiv.textContent = `✓ Document "${fileName}" generated successfully!`;
        
        // Reset form
        document.getElementById('researchForm').reset();
        
    } catch (error) {
        console.error('Error generating document:', error);
        messageDiv.className = 'message error';
        messageDiv.textContent = '✗ Error generating document. Please try again.';
    } finally {
        // Reset button
        button.disabled = false;
        button.classList.remove('loading');
        button.textContent = 'Generate Word Document';
    }
});

// Function to process and add images to document
async function addImages(files) {
    const imageParagraphs = [];
    
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        
        try {
            const arrayBuffer = await file.arrayBuffer();
            
            // Add image with caption
            imageParagraphs.push(
                new docx.Paragraph({
                    children: [
                        new docx.ImageRun({
                            data: arrayBuffer,
                            transformation: {
                                width: 500,
                                height: 375
                            }
                        })
                    ],
                    spacing: { before: 200, after: 100 },
                    alignment: docx.AlignmentType.CENTER
                }),
                new docx.Paragraph({
                    text: `Figure ${i + 1}: ${file.name}`,
                    italics: true,
                    spacing: { after: 300 },
                    alignment: docx.AlignmentType.CENTER
                })
            );
        } catch (error) {
            console.error(`Error processing image ${file.name}:`, error);
        }
    }
    
    return imageParagraphs;
}

// Optional: Add character counter
document.getElementById('content').addEventListener('input', function(e) {
    const charCount = e.target.value.length;
    let counter = document.querySelector('.char-counter');
    
    if (!counter) {
        counter = document.createElement('div');
        counter.className = 'char-counter';
        e.target.parentNode.insertBefore(counter, e.target.nextSibling);
    }
    
    counter.textContent = `${charCount} characters`;
});

