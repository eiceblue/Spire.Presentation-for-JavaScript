<template>
  <span>The example demonstrates how to set background of slide to gradient. </span>
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

        const inputFileName = "PPTSample_N.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load document from disk
        ppt.LoadFromFile(inputFileName);

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Set the background to gradient
        slide.SlideBackground.Type = wasmModule.BackgroundType.Custom;
        slide.SlideBackground.Fill.FillType = wasmModule.FillFormatType.Gradient;

        //Add gradient stops
        slide.SlideBackground.Fill.Gradient.GradientStops.Append({position:0.1,color: wasmModule.Color.get_LightSeaGreen()});
        slide.SlideBackground.Fill.Gradient.GradientStops.Append({position:0.7,color: wasmModule.Color.get_LightCyan()});

        //Set gradient shape type
        slide.SlideBackground.Fill.Gradient.GradientShape = wasmModule.GradientShapeType.Linear;

        //Set the angle
        slide.SlideBackground.Fill.Gradient.LinearGradientFill.Angle = 45;

        // Define the output file name
        const outputFileName = "SetGradientBackground_out.pptx";

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
