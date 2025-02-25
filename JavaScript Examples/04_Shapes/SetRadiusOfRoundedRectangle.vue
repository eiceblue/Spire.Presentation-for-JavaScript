<template>
  <span>The example demonstrates how to set radius of rounded rectangle. </span>
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

        //Insert a rounded rectangle and set its radious
        ppt.Slides.get_Item(0).Shapes.InsertRoundRectangle(0, 160, 180, 100, 200, 10);

        //Append a rounded rectangle and set its radius
        let shape = ppt.Slides.get_Item(0).Shapes.AppendRoundRectangle(380, 180, 100, 200, 100);
        //Set the color and fill style of shape
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_SeaGreen();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        //Rotate the shape to 90 degree
        shape.Rotation = 90;

        // Define the output file name
        const outputFileName = "SetRadiusOfRoundedRectangle_out.pptx";

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
