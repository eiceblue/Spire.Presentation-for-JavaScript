<template>
  <span>Click the following button to set text transparency in a PPT document.</span>
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
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName, rectangle:rect});
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        // Add another shape to the PPT document
        let textboxShape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(100, 100, 400, 220)});
        textboxShape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();
        textboxShape.Fill.FillType = wasmModule.FillFormatType.None;

        // Remove default blank paragraphs
        textboxShape.TextFrame.Paragraphs.Clear();

        // Add three paragraphs, apply color with different alpha values to text
        let alpha = 55;
        for (let i = 0; i < 3; i++) {
            textboxShape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());
            textboxShape.TextFrame.Paragraphs.get_Item(i).TextRanges.Append(wasmModule.TextRange.Create("Text Transparency"));
            textboxShape.TextFrame.Paragraphs.get_Item(i).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
            textboxShape.TextFrame.Paragraphs.get_Item(i).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.FromArgb({alpha,baseColor: wasmModule.Color.get_Purple()});
            alpha += 100;
        }

        const outputFileName = "SetTextTransparency.pptx";

        // Save to file
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

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
