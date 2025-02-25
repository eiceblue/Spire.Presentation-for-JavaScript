<template>
  <span>Click the following button to create Pareto chart. </span>
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

        //Create a Pareto chart in first slide
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Pareto, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

        //Set series text
        chart.ChartData._get_Item(0, 1).Text = "Series 1";

        //Set category text
        let categories = ["Category 1", "Category 2", "Category 4", "Category 3", "Category 4", "Category 2", "Category 1",
          "Category 1", "Category 3", "Category 2", "Category 4", "Category 2", "Category 3",
          "Category 1", "Category 3", "Category 2", "Category 4", "Category 1", "Category 1",
          "Category 3", "Category 2", "Category 4", "Category 1", "Category 1", "Category 3",
          "Category 2", "Category 4", "Category 1"];
        for (let i = 0; i < categories.length; i++) {
          chart.ChartData._get_Item(i + 1, 0).Text = categories[i];
        }

        //Fill data for chart
        let values = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1];
        for (let i = 0; i < values.length; i++) {
          chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
        }

        chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, 1);
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, categories.length, 0);
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, values.length, 1);
        chart.PrimaryCategoryAxis.IsBinningByCategory = true;
        chart.Series.get_Item(1).Line.FillFormat.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(1).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_Red();
        chart.ChartTitle.TextProperties.Text = "Pareto";
        chart.HasLegend = true;
        chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Bottom;

        // Define the output file name
        const outputFileName = "CreateParetoChart_out.pptx";

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
