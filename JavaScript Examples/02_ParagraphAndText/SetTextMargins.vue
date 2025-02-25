<template>
  <span>Click the following button to set margins for text inside shapes in a PPT document.</span>
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
        let ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Add a new shape to the PPT document
        let rect =  wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName, rectangle:rect});
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        // Append a new shape
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(50, 100, 500, 250)});

        // Set margins for text inside shapes
        shape.Fill.FillType = wasmModule.FillFormatType.None;
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_LightBlue();
        shape.TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Justify;
        shape.TextFrame.Text = "Using Spire.Presentation, developers will find an easy and effective method to create, read, write, modify, convert and print PowerPoint files. It's worthwhile for you to try this amazing product.";
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Rounded MT Bold");
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();

        // Set the margins for the text frame
        shape.TextFrame.MarginTop = 10;
        shape.TextFrame.MarginBottom = 35;
        shape.TextFrame.MarginLeft = 15;
        shape.TextFrame.MarginRight = 30;

        const outputFileName = "SetTextMargins.pptx";

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
