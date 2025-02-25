<template>
  <span>Click the following button to set datalabel range of charts in PowerPoint document.</span>
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

        //Create a PowerPoint document.
        let ppt = wasmModule.Presentation.Create();

        //Add a ColumnStacked chart
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ColumnStacked, rectangle: wasmModule.RectangleF.FromLTRB(100, 100, 600, 500) });

        //Set data for the chart
        let cellRange = chart.ChartData._get_ItemN("F1");
        cellRange.Text = "labelA";
        cellRange = chart.ChartData._get_ItemN("F2");
        cellRange.Text = "labelB";
        cellRange = chart.ChartData._get_ItemN("F3");
        cellRange.Text = "labelC";
        cellRange = chart.ChartData._get_ItemN("F4");
        cellRange.Text = "labelD";

        //Set data label ranges
        chart.Series.get_Item(0).DataLabelRanges = chart.ChartData._get_ItemNE("F1", "F4");

        //Add data label
        let dataLabel1 = chart.Series.get_Item(0).DataLabels.Add();
        dataLabel1.ID = 0;
        //Show the value
        dataLabel1.LabelValueVisible = false;
        //Show the label string
        dataLabel1.ShowDataLabelsRange = true;

        // Define the output file name
        const outputFileName = "SetChartDataLabelRange_out.pptx";

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
