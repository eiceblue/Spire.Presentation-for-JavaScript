<template>
  <span>Click the following button to create Waterfall chart. </span>
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

        //Create a WaterFall chart to the first slide
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.WaterFall, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

        //Set series text
        chart.ChartData._get_Item(0, 1).Text = "Series 1";

        //Set category text
        let categories = ["Category 1", "Category 2", "Category 3", "Category 4", "Category 5", "Category 6", "Category 7"];
        for (let i = 0; i < categories.length; i++) {
          chart.ChartData._get_Item(i + 1, 0).Text = categories[i];
        }

        //Fill data for chart
        let values = [100, 20, 50, -40, 130, -60, 70];
        for (let i = 0; i < values.length; i++) {
          chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
        }

        //Set series labels
        chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, 1);

        //Set categories labels
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, categories.length, 0);

        //Assign data to series values
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, values.length, 1);

        //Operate the third datapoint of first series
        let chartDataPoint = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
        chartDataPoint.Index = 2;
        chartDataPoint.SetAsTotal = true;
        chart.Series.get_Item(0).DataPoints.Add(chartDataPoint);

        //Operate the sixth datapoint of first series
        let chartDataPoint2 = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
        chartDataPoint2.Index = 5;
        chartDataPoint2.SetAsTotal = true;
        chart.Series.get_Item(0).DataPoints.Add(chartDataPoint2);
        chart.Series.get_Item(0).ShowConnectorLines = true;
        chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;

        chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Right;
        chart.ChartTitle.TextProperties.Text = "WaterFall";

        // Define the output file name
        const outputFileName = "CreateWaterFallChart_out.pptx";

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
