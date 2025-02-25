<template>
  <span>Click the following button to set line spacing.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Template_Az.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Get the first slide
        let slide = ppt.Slides.get_Item(0);

        // Create a new rectangle shape on the first slide
        let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(50, 100, ppt.SlideSize.Size.Width - 50, 400)});
        
        // Set the line color of the shape to white
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        // Set the fill type of the shape to none (transparent)
        shape.Fill.FillType = wasmModule.FillFormatType.None;

        // Clear any existing paragraphs in the text frame of the shape
        shape.TextFrame.Paragraphs.Clear();

        // Append a new text frame with specified text to the shape
        shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPointÂ® compatible API that enables developers to"
            + "create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform."
            + "From Spire.Presentation v 3.7.5, Spire.Presentation starts to support .NET Core, .NET standard.");
        
        // Retrieve the text range from the shape's text frame
        let textRange = shape.TextFrame.TextRange;

        // Set the fill type of the text to solid
        textRange.Fill.FillType = wasmModule.FillFormatType.Solid;

        // Set the solid color of the text to blue violet
        textRange.Fill.SolidColor.Color = wasmModule.Color.get_BlueViolet();

        // Set the font height of the text to 20 points
        textRange.FontHeight = 20;

        // Set the Latin font of the text to "Lucida Sans Unicode"
        textRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

        // Adjust paragraph spacing for the first paragraph in the text frame
        shape.TextFrame.Paragraphs.get_Item(0).SpaceBefore = 100;
        shape.TextFrame.Paragraphs.get_Item(0).SpaceAfter = 100;
        shape.TextFrame.Paragraphs.get_Item(0).LineSpacing = 150;

        const outputFileName = "LineSpacing.pptx";

        // Save to file
        ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

        // Clean up resources
        ppt.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
