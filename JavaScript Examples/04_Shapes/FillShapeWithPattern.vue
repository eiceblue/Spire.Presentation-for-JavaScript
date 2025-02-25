<template>
  <span>The example demonstrates how to fill shape with pattern in a PPT document. .</span>
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

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Add a rectangle
        let rect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 50, 100, (100 + ppt.SlideSize.Size.Width / 2 - 50), 200);
        let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:rect});

        //Set the pattern fill format
        shape.Fill.FillType = wasmModule.FillFormatType.Pattern;
        shape.Fill.Pattern.PatternType = wasmModule.PatternFillType.Trellis;
        shape.Fill.Pattern.BackgroundColor.Color = wasmModule.Color.get_DarkGray();
        shape.Fill.Pattern.ForegroundColor.Color = wasmModule.Color.get_Yellow();

        //Set the fill format of line
        shape.Line.FillType = wasmModule.FillFormatType.Solid;
        shape.Line.SolidFillColor.Color = wasmModule.Color.get_Transparent();

        // Define the output file name
        const outputFileName = "FillShapeWithPattern_out.pptx";

        // Save the document to the specified path
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
