<template>
  <span>The following example demonstrates how to set tick-mark labels on the category axis</span>
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
      
        // Load the input file into the virtual file system (VFS)
        const inputFileName = "Template_Ppt_3.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PowerPonit document
        let ppt = wasmModule.Presentation.Create();

        // Load the file from disk
        ppt.LoadFromFile(inputFileName);

        // Get the chart from the PowerPoint slide
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        // Rotate tick labels
        chart.PrimaryCategoryAxis.TextRotationAngle = 45;

        // Specify interval between labels
        chart.PrimaryCategoryAxis.IsAutomaticTickLabelSpacing = false;
        chart.PrimaryCategoryAxis.TickLabelSpacing = 2;

        // Change position
        chart.PrimaryCategoryAxis.TickLabelPosition = wasmModule.TickLabelPositionType.TickLabelPositionHigh;

        // Define the output file name
        const outputFileName = "SetTickMarkLabelsOnCategoryAxis.pptx";

        // Save the document 
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
