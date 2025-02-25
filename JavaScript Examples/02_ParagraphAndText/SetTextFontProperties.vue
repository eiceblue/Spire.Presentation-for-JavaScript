<template>
  <span>Click the following button to set text font properties.</span>
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

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Add a new shape to the PPT document
        let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 250, 80, (500 + ppt.SlideSize.Size.Width / 2 - 250), 230);
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rec});

        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
        shape.Fill.FillType = wasmModule.FillFormatType.None;

        // Add text to the shape
        shape.AppendTextFrame("Welcome to use Spire.Presentation");

        let textRange = shape.TextFrame.TextRange;

        // Set the font
        textRange.LatinFont = wasmModule.TextFont.Create("Times New Roman");

        // Set bold property of the font
        textRange.IsBold = wasmModule.TriState.True;

        // Set italic property of the font
        textRange.IsItalic = wasmModule.TriState.True;

        // Set underline property of the font
        textRange.TextUnderlineType = wasmModule.TextUnderlineType.Single;

        // Set the height of the font
        textRange.FontHeight = 50;

        // Set the color of the font
        textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        textRange.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();

        const outputFileName = "SetTextFontProperties.pptx";

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
