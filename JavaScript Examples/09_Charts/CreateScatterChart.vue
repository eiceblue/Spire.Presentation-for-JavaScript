<template>
  <span>Click the following button to create scatterd chart. </span>
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

        //Creat a presentation
        let ppt = wasmModule.Presentation.Create();

        //Set background image
        let rect2 = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: inputFileName, rectangle: rect2 });
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Insert a chart and set chart title and chart type
        let rect1 = wasmModule.RectangleF.FromLTRB(90, 100, 640, 420);
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ScatterMarkers, rectangle: rect1, init: false });
        chart.ChartTitle.TextProperties.Text = "ScatterMarker Chart";
        chart.ChartTitle.TextProperties.IsCentered = true;
        chart.ChartTitle.Height = 30;
        chart.HasTitle = true;

        //Set chart data
        let xdata = [2.7, 8.9, 10.0, 12.4];
        let ydata = [3.2, 15.3, 6.7, 8];

        chart.ChartData._get_Item(0, 0).Text = "X-Value";
        chart.ChartData._get_Item(0, 1).Text = "Y-Value";

        for (let i = 0; i < xdata.length; ++i) {
          chart.ChartData._get_Item(i + 1, 0).NumberValue = xdata[i];
          chart.ChartData._get_Item(i + 1, 1).NumberValue = ydata[i];
        }

        //Set the series label
        chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "B1");

        //Assign data to X axis, Y axis and Bubbles
        chart.Series.get_Item(0).XValues = chart.ChartData._get_ItemNE("A2", "A5");
        chart.Series.get_Item(0).YValues = chart.ChartData._get_ItemNE("B2", "B5");

        // Define the output file name
        const outputFileName = "CreateScatterChart_out.pptx";

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
