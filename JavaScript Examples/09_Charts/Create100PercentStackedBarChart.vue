<template>
  <span>Click the following button to create 100% stacked bar chart. </span>
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

        //Add a "Bar100PercentStacked" chart to the first slide.
        ppt.SlideSize.Type = wasmModule.SlideSizeType.Screen16x9;
        let slidesize = ppt.SlideSize.Size;

        let slide = ppt.Slides.get_Item(0);

        //Append a chart.
        let rect = wasmModule.RectangleF.FromLTRB(20, 20, slidesize.Width - 20, slidesize.Height - 20);
        let chart = slide.Shapes.AppendChart({ type: wasmModule.ChartType.Bar100PercentStacked, rectangle: rect });

        //Write data to the chart data.
        let columnlabels = ["Series 1", "Series 2", "Series 3"];

        //Insert the column labels.
        for (let i = 0; i < columnlabels.length; i++) {
          chart.ChartData._get_Item(0, i + 1).Text = columnlabels[i];
        }

        let rowlabels = ["Category 1", "Category 2", "Category 3"];

        //Insert the row labels.
        for (let i = 0; i < rowlabels.length; i++) {
          chart.ChartData._get_Item(i + 1, 0).Text = rowlabels[i];
        }

        let values = [[20.83233, 10.34323, -10.354667], [10.23456, -12.23456, 23.34456], [12.34345, -23.34343, -13.23232]];

        //Insert the values.
        let value = 0.0;
        for (let i = 0; i < rowlabels.length; i++) {
          for (let j = 0; j < columnlabels.length; j++) {
            value = Math.round(values[i][j], 2);
            chart.ChartData._get_Item(i + 1, j + 1).NumberValue = value;
          }
        }

        chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, columnlabels.length);
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, rowlabels.length, 0);

        //Set the position of category axis.
        chart.PrimaryCategoryAxis.Position = wasmModule.AxisPositionType.Left;
        chart.SecondaryCategoryAxis.Position = wasmModule.AxisPositionType.Left;
        chart.PrimaryCategoryAxis.TickLabelPosition = wasmModule.TickLabelPositionType.TickLabelPositionLow;

        //Set the data, font and format for the series of each column.
        for (let i = 0; i < columnlabels.length; i++) {
          chart.Series.get_Item(i).Values = chart.ChartData._get_ItemRCLL(1, i + 1, rowlabels.length, i + 1);
          chart.Series.get_Item(i).Fill.FillType = wasmModule.FillFormatType.Solid;
          chart.Series.get_Item(i).InvertIfNegative = false;
          for (let j = 0; j < rowlabels.length; j++) {
            let label = chart.Series.get_Item(i).DataLabels.Add();
            label.LabelValueVisible = true;
            chart.Series.get_Item(i).DataLabels.get_Item(j).HasDataSource = false;
            chart.Series.get_Item(i).DataLabels.get_Item(j).NumberFormat = "0#\\%";
            chart.Series.get_Item(i).DataLabels.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 12;
          }
        }

        //Set the color of the Series.
        chart.Series.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_YellowGreen();
        chart.Series.get_Item(1).Fill.SolidColor.Color = wasmModule.Color.get_Red();
        chart.Series.get_Item(2).Fill.SolidColor.Color = wasmModule.Color.get_Green();

        let font = wasmModule.TextFont.Create("Tw Cen MT");

        //Set the font and size for chartlegend.
        for (let k = 0; k < chart.ChartLegend.EntryTextProperties.length; k++) {
          let textPara = chart.ChartLegend.EntryTextProperties[k];
          textPara.LatinFont = font;
          textPara.FontHeight = 20;
        }

        // Define the output file name
        const outputFileName = "Create100PercentStackedBarChart_out.pptx";

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
