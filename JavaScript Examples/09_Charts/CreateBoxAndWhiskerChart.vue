<template>
  <span>Click the following button to create BoxAndWhiske chart. </span>
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

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Insert a BoxAndWhisker chart to the first slide
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.BoxAndWhisker, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

        // Series labels
        let seriesLabel = ["Series 1", "Series 2", "Series 3"];
        for (let i = 0; i < seriesLabel.length; i++) {
          chart.ChartData._get_Item(0, i + 1).Text = "Series 1";
        }

        // Categories
        let categories = ["Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1",
          "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2",
          "Category 3", "Category 3", "Category 3", "Category 3", "Category 3"];
        for (let i = 0; i < categories.length; i++) {
          chart.ChartData._get_Item(i + 1, 0).Text = categories[i];
        }

        // Values
        let values = [[-7, -3, -24], [-10, 1, 11], [-28, -6, 34], [47, 2, -21], [35, 17, 22], [-22, 15, 19], [17, -11, 25],
        [-30, 18, 25], [49, 22, 56], [37, 22, 15], [-55, 25, 31], [14, 18, 22], [18, -22, 36], [-45, 25, -17],
        [-33, 18, 22], [18, 2, -23], [-33, -22, 10], [10, 19, 22]];
        for (let i = 0; i < seriesLabel.length; i++) {
          for (let j = 0; j < categories.length; j++) {
            chart.ChartData._get_Item(j + 1, i + 1).NumberValue = values[j][i];
          }
        }

        //Set series
        chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, seriesLabel.length);
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, categories.length, 0);
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, categories.length, 1);
        chart.Series.get_Item(1).Values = chart.ChartData._get_ItemRCLL(1, 2, categories.length, 2);
        chart.Series.get_Item(2).Values = chart.ChartData._get_ItemRCLL(1, 3, categories.length, 3);
        chart.Series.get_Item(0).ShowInnerPoints = false;
        chart.Series.get_Item(0).ShowOutlierPoints = true;
        chart.Series.get_Item(0).ShowMeanMarkers = true;
        chart.Series.get_Item(0).ShowMeanLine = true;
        chart.Series.get_Item(0).QuartileCalculationType = wasmModule.QuartileCalculation.ExclusiveMedian;
        chart.Series.get_Item(1).ShowInnerPoints = false;
        chart.Series.get_Item(1).ShowOutlierPoints = true;
        chart.Series.get_Item(1).ShowMeanMarkers = true;
        chart.Series.get_Item(1).ShowMeanLine = true;
        chart.Series.get_Item(1).QuartileCalculationType = wasmModule.QuartileCalculation.InclusiveMedian;
        chart.Series.get_Item(2).ShowInnerPoints = false;
        chart.Series.get_Item(2).ShowOutlierPoints = true;
        chart.Series.get_Item(2).ShowMeanMarkers = true;
        chart.Series.get_Item(2).ShowMeanLine = true;
        chart.Series.get_Item(2).QuartileCalculationType = wasmModule.QuartileCalculation.ExclusiveMedian;

        //Show legend
        chart.HasLegend = true;
        chart.ChartTitle.TextProperties.Text = "BoxAndWhisker";
        chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Top;

        // Define the output file name
        const outputFileName = "CreateBoxAndWhiskerChart_out.pptx";

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
