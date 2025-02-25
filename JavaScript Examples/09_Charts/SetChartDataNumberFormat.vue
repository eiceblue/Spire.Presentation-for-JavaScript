<template>
  <span>Click the following button to set the number format for chart data.</span>
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
        let inputFileName = "SetChartDataNumberFormat.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);


        //Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get chart on the first slide
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Set the number format for Axis
        chart.PrimaryValueAxis.NumberFormat = "#,##0.00";

        //Set the DataLabels format for Axis
        chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;
        chart.Series.get_Item(0).DataLabels.PercentValueVisible = false;
        chart.Series.get_Item(0).DataLabels.NumberFormat = "#,##0.00";
        chart.Series.get_Item(0).DataLabels.HasDataSource = false;

        //Set the number format for ChartData
        for (let i = 1; i <= chart.Series.get_Item(0).Values.Count; i++) {
          chart.ChartData._get_Item(i, 1).NumberFormat = "#,##0.00";
        }

        // Define the output file name
        const outputFileName = "SetChartDataNumberFormat_out.pptx";

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
