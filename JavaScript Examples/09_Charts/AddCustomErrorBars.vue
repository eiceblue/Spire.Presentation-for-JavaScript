<template>
  <span>Click the following button to add custom error bars for chart in a PPT document.</span>
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
        let inputFileName = 'ChartSample1.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get the bubble chart on the first slide
        let bubbleChart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Get X error bars of the first chart series
        let errorBarsXFormat = bubbleChart.Series.get_Item(0).ErrorBarsXFormat;
        //Specify error amount type as custom error bars
        errorBarsXFormat.ErrorBarvType = wasmModule.ErrorValueType.CustomErrorBars;
        //Set the minus and plus value of the X error bars
        errorBarsXFormat.MinusVal = 0.5;
        errorBarsXFormat.PlusVal = 0.5;

        //Get Y error bars of the first chart series
        let errorBarsYFormat = bubbleChart.Series.get_Item(0).ErrorBarsYFormat;
        //Specify error amount type as custom error bars
        errorBarsYFormat.ErrorBarvType = wasmModule.ErrorValueType.CustomErrorBars;
        //Set the minus and plus value of the Y error bars
        errorBarsYFormat.MinusVal = 1;
        errorBarsYFormat.PlusVal = 1;

        // Define the output file name
        const outputFileName = "AddCustomErrorBars_out.pptx";

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
