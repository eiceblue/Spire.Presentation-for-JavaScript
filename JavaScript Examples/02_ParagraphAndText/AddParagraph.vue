<template>
  <span>Click the following button to add a paragraph to PPT document.</span>
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

        // Load the image into the virtual file system (VFS)
        let imageName = "bg.png";
        await wasmModule.FetchFileToVFS(imageName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Create a rectangle from the left, top, right, and bottom coordinates based on slide size
        let rec = wasmModule.RectangleF.FromLTRB(0,0,ppt.SlideSize.Size.Width,ppt.SlideSize.Size.Height);

        // Append an embedded image as a rectangle shape to the first slide
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName:imageName,rectangle:rec});

        // Set the line color of the first shape to Floral White
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        // Append a new shape
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(50, 70, 670, 220)});
        shape.Fill.FillType = wasmModule.FillFormatType.None;
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        // Set the alignment of paragraph
        shape.TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Left;

        // Set the indent of paragraph
        shape.TextFrame.Paragraphs.get_Item(0).Indent = 50;

        // Set the linespacing of paragraph
        shape.TextFrame.Paragraphs.get_Item(0).LineSpacing = 150;

        // Set the text of paragraph
        shape.TextFrame.Text = "This powerful component suite contains the most up-to-date versions of all .NET WPF Silverlight components offered by E-iceblue.";

        // Set the Font
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Rounded MT Bold");
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();

        // Define the output file name 
        const outputFileName = "AddParagraph.pptx";

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
