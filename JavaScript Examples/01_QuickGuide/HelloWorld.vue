<template>
  <span>Click the following button to write "Hello World" into a PPT document.</span>
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
        let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 250,80,(500 + ppt.SlideSize.Size.Width / 2 - 250),230);
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:rec});

        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
        shape.Fill.FillType = wasmModule.FillFormatType.None;

        // Add text to the shape
        shape.AppendTextFrame("Hello World!");

        // Set the font and fill style of the text
        let textRange = shape.TextFrame.TextRange;
        textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        textRange.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
        textRange.FontHeight = 66;
        textRange.LatinFont = wasmModule.TextFont;

        // Define the output file name 
        const outputFileName = "HelloWorld.pptx";

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
