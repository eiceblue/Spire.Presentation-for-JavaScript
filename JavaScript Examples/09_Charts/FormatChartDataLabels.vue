<template>
  <span>Click the following button to format chart dataLabels.</span>
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
        let inputFileName = "FormatChartDataLabels.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create PPT document and load file.
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get the chart
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Get the chart series
        let sers = chart.Series;

        //Initialize four instances of series label and set parameters of each label
        let cd1 = sers.get_Item(0).DataLabels.Add();
        cd1.PercentageVisible = true;
        cd1.TextFrame.Text = "Custom Datalabel1";
        cd1.TextFrame.TextRange.FontHeight = 12;
        cd1.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
        cd1.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        cd1.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_Green();

        let cd2 = sers.get_Item(0).DataLabels.Add();
        cd2.Position = wasmModule.ChartDataLabelPosition.InsideEnd;
        cd2.PercentageVisible = true;
        cd2.TextFrame.Text = "Custom Datalabel2";
        cd2.TextFrame.TextRange.FontHeight = 10;
        cd2.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Arial");
        cd2.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        cd2.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_OrangeRed();

        let cd3 = sers.get_Item(0).DataLabels.Add();
        cd3.Position = wasmModule.ChartDataLabelPosition.Center;
        cd3.PercentageVisible = true;
        cd3.TextFrame.Text = "Custom Datalabel3";
        cd3.TextFrame.TextRange.FontHeight = 14;
        cd3.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Calibri");
        cd3.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        cd3.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_Blue();

        let cd4 = sers.get_Item(0).DataLabels.Add();
        cd4.Position = wasmModule.ChartDataLabelPosition.InsideBase;
        cd4.PercentageVisible = true;
        cd4.TextFrame.Text = "Custom Datalabel4";
        cd4.TextFrame.TextRange.FontHeight = 12;
        cd4.TextFrame.TextRange.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
        cd4.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        cd4.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_OliveDrab();

        // Define the output file name
        const outputFileName = "FormatChartDataLabels_out.pptx";

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
