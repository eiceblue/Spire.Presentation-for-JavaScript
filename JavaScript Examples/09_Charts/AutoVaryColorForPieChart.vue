<template>
  <span>Click the following button to auto vary color for pie chart. </span>
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

        let rect1 = wasmModule.RectangleF.FromLTRB(40, 100, 590, 420);
        //Add a pie chart
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Pie, rectangle: rect1, init: false });
        chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
        chart.ChartTitle.TextProperties.IsCentered = true;
        chart.ChartTitle.Height = 30;
        chart.HasTitle = true;
        //Attach the data to chart
        let quarters = ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"];
        let sales = [210, 320, 180, 500];
        chart.ChartData._get_Item(0, 0).Text = "Quarters";
        chart.ChartData._get_Item(0, 1).Text = "Sales";
        for (let i = 0; i < quarters.length; ++i) {
          chart.ChartData._get_Item(i + 1, 0).Text = quarters[i];
          chart.ChartData._get_Item(i + 1, 1).NumberValue = sales[i];
        }

        chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A5");
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B5");

        //Set whether auto vary color, default value is true
        chart.Series.get_Item(0).IsVaryColor = false;
        chart.Series.get_Item(0).Distance = 15;

        // Define the output file name
        const outputFileName = "AutoVaryColorForPieChart_out.pptx";

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
