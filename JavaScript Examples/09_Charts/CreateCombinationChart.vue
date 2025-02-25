<template>
  <span>Click the following button to create a combination chart. </span>
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

        //Create a presentation instance
        let ppt = wasmModule.Presentation.Create();

        //Set background image
        let rect2 = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: inputFileName, rectangle: rect2 });
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Insert a column clustered chart
        let rect = wasmModule.RectangleF.FromLTRB(100, 100, 650, 420);
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ColumnClustered, rectangle: rect });

        //Set chart title
        chart.ChartTitle.TextProperties.Text = "Monthly Sales Report";
        chart.ChartTitle.TextProperties.IsCentered = true;
        chart.ChartTitle.Height = 30;
        chart.HasTitle = true;

        //Create a datatable
        let caption = ["Month", "Sales", "Growth rate"];
        let month = ["January", "February", "March", "April", "May", "June"];
        let sales = [200, 250, 300, 150, 200, 400];
        let growth_rate = [0.6, 0.8, 0.6, 0.2, 0.5, 0.9];

        //Import data from datatable to chart data
        for (let i = 0; i < caption.length; i++) {
          chart.ChartData._get_Item(0, i).Text = caption[i];
        }
        for (let i = 0; i < month.length; i++) {
          chart.ChartData._get_Item(i + 1, 0).Text = month[i];
        }
        for (let i = 0; i < sales.length; i++) {
          chart.ChartData._get_Item(i + 1, 1).NumberValue = sales[i];
        }
        for (let i = 0; i < growth_rate.length; i++) {
          chart.ChartData._get_Item(i + 1, 2).NumberValue = growth_rate[i];
        }


        //Set series labels
        chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "C1");

        //Set categories labels
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A7");

        //Assign data to series values
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B7");
        chart.Series.get_Item(1).Values = chart.ChartData._get_ItemNE("C2", "C7");

        //Change the chart type of serie 2 to line with markers
        chart.Series.get_Item(1).Type = wasmModule.ChartType.LineMarkers;

        //Plot data of series 2 on the secondary axis
        chart.Series.get_Item(1).UseSecondAxis = true;

        //Set the number format as percentage
        chart.SecondaryValueAxis.NumberFormat = "0%";

        //Hide gridlinkes of secondary axis
        chart.SecondaryValueAxis.MajorGridTextLines.FillType = wasmModule.FillFormatType.None;

        //Set overlap
        chart.OverLap = -50;

        //Set gapwidth
        chart.GapWidth = 200;


        // Define the output file name
        const outputFileName = "CreateCombinationChart_out.pptx";

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
