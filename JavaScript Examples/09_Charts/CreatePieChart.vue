<template>
  <span>Click the following button to create pie chart in PowerPoint document. </span>
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

        //Insert a Pie chart to the first slide and set the chart title.
        let rect1 = wasmModule.RectangleF.FromLTRB(40, 100, 590, 420);
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Pie, rectangle: rect1, init: false });
        chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
        chart.ChartTitle.TextProperties.IsCentered = true;
        chart.ChartTitle.Height = 30;
        chart.HasTitle = true;

        //Define some data.
        let quarters = ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"];
        let sales = [210, 320, 180, 500];

        //Append data to ChartData, which represents a data table where the chart data is stored.
        chart.ChartData._get_Item(0, 0).Text = "Quarters";
        chart.ChartData._get_Item(0, 1).Text = "Sales";
        for (let i = 0; i < quarters.length; ++i) {
          chart.ChartData._get_Item(i + 1, 0).Text = quarters[i];
          chart.ChartData._get_Item(i + 1, 1).NumberValue = sales[i];
        }

        //Set category labels, series label and series data.
        chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A5");
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B5");

        //Add data points to series and fill each data point with different color.
        for (let i = 0; i < chart.Series.get_Item(0).Values.Count; i++) {
          let cdp = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
          cdp.Index = i;
          chart.Series.get_Item(0).DataPoints.Add(cdp);
        }
        chart.Series.get_Item(0).DataPoints.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).DataPoints.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_RosyBrown();
        chart.Series.get_Item(0).DataPoints.get_Item(1).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).DataPoints.get_Item(1).Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
        chart.Series.get_Item(0).DataPoints.get_Item(2).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).DataPoints.get_Item(2).Fill.SolidColor.Color = wasmModule.Color.get_LightPink();
        chart.Series.get_Item(0).DataPoints.get_Item(3).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).DataPoints.get_Item(3).Fill.SolidColor.Color = wasmModule.Color.get_MediumPurple();

        //Set the data labels to display label value and percentage value.
        chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;
        chart.Series.get_Item(0).DataLabels.PercentValueVisible = true;

        // Define the output file name
        const outputFileName = "CreatePieChart_out.pptx";

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
