<template>
  <span>Click the following button to add and format error bars of charts.</span>
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
        let inputFileName = 'AddAndFormatErrorBars.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PowerPoint document.
        let ppt = wasmModule.Presentation.Create();

        //Load the file.
        ppt.LoadFromFile(inputFileName);

        //Get the column chart on the first slide and set chart title.
        let columnChart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        columnChart.ChartTitle.TextProperties.Text = "Vertical Error Bars";

        //Add Y (Vertical) Error Bars.

        //Get Y error bars of the first chart series.
        let errorBarsYFormat1 = columnChart.Series.get_Item(0).ErrorBarsYFormat;

        //Set end cap.
        errorBarsYFormat1.ErrorBarNoEndCap = false;

        //Specify direction.
        errorBarsYFormat1.ErrorBarSimType = wasmModule.ErrorBarSimpleType.Plus;

        //Specify error amount type.
        errorBarsYFormat1.ErrorBarvType = wasmModule.ErrorValueType.StandardError;

        //Set value.
        errorBarsYFormat1.ErrorBarVal = 0.3;

        //Set line format.
        errorBarsYFormat1.Line.FillType = wasmModule.FillFormatType.Solid;
        errorBarsYFormat1.Line.SolidFillColor.Color = wasmModule.Color.get_MediumVioletRed();
        errorBarsYFormat1.Line.Width = 1;

        //Get the bubble chart on the second slide and set chart title.
        let bubbleChart = ppt.Slides.get_Item(1).Shapes.get_Item(0);
        bubbleChart.ChartTitle.TextProperties.Text = "Vertical and Horizontal Error Bars";

        //Add X (Horizontal) and Y (Vertical) Error Bars.
        //Get X error bars of the first chart series.
        let errorBarsXFormat = bubbleChart.Series.get_Item(0).ErrorBarsXFormat;

        //Set end cap.
        errorBarsXFormat.ErrorBarNoEndCap = false;

        //Specify direction.
        errorBarsXFormat.ErrorBarSimType = wasmModule.ErrorBarSimpleType.Both;

        //Specify error amount type.
        errorBarsXFormat.ErrorBarvType = wasmModule.ErrorValueType.StandardError;

        //Set value.
        errorBarsXFormat.ErrorBarVal = 0.3;

        //Get Y error bars of the first chart series.
        let errorBarsYFormat2 = bubbleChart.Series.get_Item(0).ErrorBarsYFormat;

        //Set end cap.
        errorBarsYFormat2.ErrorBarNoEndCap = false;

        //Specify direction.
        errorBarsYFormat2.ErrorBarSimType = wasmModule.ErrorBarSimpleType.Both;

        //Specify error amount type.
        errorBarsYFormat2.ErrorBarvType = wasmModule.ErrorValueType.StandardError;

        //Set value.
        errorBarsYFormat2.ErrorBarVal = 0.3;

        // Define the output file name
        const outputFileName = "AddAndFormatErrorBars_out.pptx";

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
