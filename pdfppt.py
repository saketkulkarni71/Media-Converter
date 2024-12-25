import aspose.slides as slides

# Load presentation
pres = slides.Presentation("presentation.pptx")

# Convert PPTX to PDF
pres.save("hello.pdf", slides.export.SaveFormat.PDF)