<template>
  <span>Click the following button to change text font in chart. </span>
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
        let inputFileName = 'ChangeTextFontInChart.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create Presentation
        let ppt = wasmModule.Presentation.Create();

        //Load a PPTX file
        ppt.LoadFromFile(inputFileName);

        //Get the chart
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Change the font of title
        chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
        chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Blue;
        chart.ChartTitle.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 30;

        //Change the font of legend
        chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.DarkGreen;
        chart.ChartLegend.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

        //Change the font of series
        chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Red;
        chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.FontHeight = 10;
        chart.PrimaryCategoryAxis.TextProperties.Paragraphs.get_Item(0).DefaultCharacterProperties.LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");

        // Define the output file name
        const outputFileName = "ChangeTextFontInChart_out.pptx";

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
