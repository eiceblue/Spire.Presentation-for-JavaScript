<template>
  <span>The example shows how to add lines to slide with two points in a PPT document.</span>
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

        //Add line with two points
        let line = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line,start: wasmModule.PointF.Create(50, 50),end: wasmModule.PointF.Create(150, 150)});
        line.ShapeStyle.LineColor.Color = wasmModule.Color.get_Red();
        line = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line,start: wasmModule.PointF.Create(150, 150),end: wasmModule.PointF.Create(250, 50)});
        line.ShapeStyle.LineColor.Color = wasmModule.Color.get_Blue();

        // Define the output file name
        const outputFileName = "AddLineWithTwoPoints_out.pptx";

        // Save the document to the specified path
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013});

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
