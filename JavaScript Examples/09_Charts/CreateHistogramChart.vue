<template>
  <span>Click the following button to create Histogram chart. </span>
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

        //Create PPT document
        let ppt = wasmModule.Presentation.Create();

        //Add a Histogram chart
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Histogram, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

        //Set series text
        chart.ChartData._get_Item(0, 0).Text = "Series 1";

        //Fill data for chart
        let values = [1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 8, 8, 8, 9, 9, 9, 12, 12, 13, 13, 17, 17, 17, 19, 19, 19, 25, 25, 25, 25, 25, 25, 25, 25, 29, 29, 29, 29, 32, 32, 33, 33, 35, 35, 41, 41, 44, 45, 49, 49];
        for (let i = 0; i < values.length; i++) {
          chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
        }

        //Set series label
        chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 0, 0, 0);

        //Set values for series
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 0, values.length, 0);

        chart.PrimaryCategoryAxis.NumberOfBins = 7;
        chart.PrimaryCategoryAxis.GapWidth = 20;
        //Chart title
        chart.ChartTitle.TextProperties.Text = "Histogram";
        chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Bottom;

        // Define the output file name
        const outputFileName = "CreateHistogramChart_out.pptx";

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
