<template>
  <span>Click the following button to create doughnut chart. </span>
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
        let inputFileName = 'bg.png';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a ppt document
        let ppt = wasmModule.Presentation.Create();

        let rect = wasmModule.RectangleF.FromLTRB(80, 100, 630, 420);

        //Set background image
        let rect2 = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: inputFileName, rectangle: rect2 });
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Add a Doughnut chart
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Doughnut, rectangle: rect, init: false });
        chart.ChartTitle.TextProperties.Text = "Market share by country";
        chart.ChartTitle.TextProperties.IsCentered = true;
        chart.ChartTitle.Height = 30;

        let countries = ["Guba", "Mexico", "France", "German"];
        let sales = [1800, 3000, 5100, 6200];
        chart.ChartData._get_Item(0, 0).Text = "Countries";
        chart.ChartData._get_Item(0, 1).Text = "Sales";
        for (let i = 0; i < countries.length; ++i) {
          chart.ChartData._get_Item(i + 1, 0).Text = countries[i];
          chart.ChartData._get_Item(i + 1, 1).NumberValue = sales[i];
        }
        chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A5");
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B5");

        for (let i = 0; i < chart.Series.get_Item(0).Values.Count; i++) {
          let cdp = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
          cdp.Index = i;
          chart.Series.get_Item(0).DataPoints.Add(cdp);
        }
        //Set the series color
        chart.Series.get_Item(0).DataPoints.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).DataPoints.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
        chart.Series.get_Item(0).DataPoints.get_Item(1).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).DataPoints.get_Item(1).Fill.SolidColor.Color = wasmModule.Color.get_MediumPurple();
        chart.Series.get_Item(0).DataPoints.get_Item(2).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).DataPoints.get_Item(2).Fill.SolidColor.Color = wasmModule.Color.get_DarkGray();
        chart.Series.get_Item(0).DataPoints.get_Item(3).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).DataPoints.get_Item(3).Fill.SolidColor.Color = wasmModule.Color.get_DarkOrange();

        chart.Series.get_Item(0).DataLabels.LabelValueVisible = true;
        chart.Series.get_Item(0).DataLabels.PercentValueVisible = true;
        chart.Series.get_Item(0).DoughnutHoleSize = 60;


        // Define the output file name
        const outputFileName = "CreateDoughnutChart_out.pptx";

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
