<template>
  <span>The example demonstrates how to apply formatting on rectangle in a PPT document. </span>
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

        //Add a shape
        let rect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 100, 100, (200 + ppt.SlideSize.Size.Width / 2 - 100), 200);
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rect});

        //Set the fill format of shape
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();

        //Set the fill format of line
        shape.Line.FillType = wasmModule.FillFormatType.Solid;
        shape.Line.SolidFillColor.Color = wasmModule.Color.get_DimGray();

        // Define the output file name
        const outputFileName = "SetRectangleFormat_out.pptx";

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
