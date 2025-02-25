<template>
  <span>Click the following button to set custom fonts.</span>
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

        // Create a rectangle using the specified coordinates
        let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 250, 80, (500 + ppt.SlideSize.Size.Width / 2 - 250), 230);
        
        // Append a rectangle shape to the first slide using the defined rectangle
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:rec});

        // Set the line color of the rectangle shape to white
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        // Set the fill type of the rectangle shape to none, making it transparent
        shape.Fill.FillType = wasmModule.FillFormatType.None;

        // Add text to the shape
        shape.AppendTextFrame("Hello World!");

        // Set the custom font folder
        ppt.SetCustomFontsFolder("/Library/Fonts/");

        // Set the font and fill style of the text
        let textRange = shape.TextFrame.TextRange;
        textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        textRange.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
        textRange.FontHeight = 66;
        textRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

        const outputFileName = "SetCustomFonts.pptx";

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
