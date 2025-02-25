<template>
  <span>Click the following button to set datapoint color in chart.</span>
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
        let inputFileName = "SetDatapointColorInChart.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get the chart
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Initialize an instances of dataPoint
        let cdp1 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));

        //Specify the datapoint order
        cdp1.Index = 0;

        //Set the color of the datapoint
        cdp1.Fill.FillType = wasmModule.FillFormatType.Solid;
        cdp1.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Orange;

        //Add the dataPoint to first series
        chart.Series.get_Item(0).DataPoints.Add(cdp1);

        //Set the color for the other three data points
        let cdp2 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
        cdp2.Index = 1;
        cdp2.Fill.FillType = wasmModule.FillFormatType.Solid;
        cdp2.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Gold;
        chart.Series.get_Item(0).DataPoints.Add(cdp2);

        let cdp3 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
        cdp3.Index = 2;
        cdp3.Fill.FillType = wasmModule.FillFormatType.Solid;
        cdp3.Fill.SolidColor.KnownColor = wasmModule.KnownColors.MediumPurple;
        chart.Series.get_Item(0).DataPoints.Add(cdp3);

        let cdp4 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
        cdp4.Index = 1;
        cdp4.Fill.FillType = wasmModule.FillFormatType.Solid;
        cdp4.Fill.SolidColor.KnownColor = wasmModule.KnownColors.ForestGreen;
        chart.Series.get_Item(0).DataPoints.Add(cdp4);

        // Define the output file name
        const outputFileName = "SetDatapointColorInChart_out.pptx";

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
