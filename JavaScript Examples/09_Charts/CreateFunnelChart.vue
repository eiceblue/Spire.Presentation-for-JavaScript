<template>
  <span>Click the following button to create Funnel chart. </span>
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

        //Create a Funnel chart to the first slide
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Funnel, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 600, 450), init: false });

        //Set series text
        chart.ChartData._get_Item(0, 1).Text = "Series 1";

        //Set category text
        let categories = ["Website Visits", "Download", "Uploads", "Requested price", "Invoice sent", "Finalized"];
        for (let i = 0; i < categories.length; i++) {
          chart.ChartData._get_Item(i + 1, 0).Text = categories[i];
        }

        //Fill data for chart
        let values = [50000, 47000, 30000, 15000, 9000, 5600];
        for (let i = 0; i < values.length; i++) {
          chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
        }

        //Set series labels
        chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, 1);

        //Set categories labels
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, categories.length, 0);

        //Assign data to series values
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, values.length, 1);

        //Set the chart title
        chart.ChartTitle.TextProperties.Text = "Funnel";


        // Define the output file name
        const outputFileName = "CreateFunnelChart_out.pptx";

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
