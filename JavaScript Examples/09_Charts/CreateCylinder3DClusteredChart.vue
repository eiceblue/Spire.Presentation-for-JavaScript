<template>
  <span>Click the following button to create Cylinder3Dclustered chart into a PPT document. </span>
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
    // aconst fs = require('fs');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {

        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'bg.png';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = 'data.xml';
        await wasmModule.FetchFileToVFS(inputFileName2, '', `${import.meta.env.BASE_URL}static/data/`);

        // Read XML content from the input file
        const data = wasmModule.FS.readFile(inputFileName2);
        
        // Create a TextDecoder instance to decode the Uint8Array into a string using UTF-8 encoding
        const decoder = new TextDecoder('utf-8');
        
        // Decode the Uint8Array data into a string
        const stringData = decoder.decode(data);

        // Parse the decoded string as XML using DOMParser
        const xmlDocument = new DOMParser().parseFromString(stringData,'application/xml');

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Set background image
        let rect2 = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: inputFileName, rectangle: rect2 });
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Insert chart
        let rect = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 200, 85, (400 + ppt.SlideSize.Size.Width / 2 - 200), 485);
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Cylinder3DClustered, rectangle: rect });

        //Add chart Title
        chart.ChartTitle.TextProperties.Text = "Report";
        chart.ChartTitle.TextProperties.IsCentered = true;
        chart.ChartTitle.Height = 30;
        chart.HasTitle = true;

        //Load data from XML file to datatable
        let caption = ["SalesPers", "SaleAmt", "ComPct", "ComAmt"];
        for (let i = 0; i < caption.length; i++) {
          chart.ChartData._get_Item(0, i).Text = caption[i];
        }
        for (let i = 0; i < xmlDocument.documentElement.getElementsByTagName('Report').length; i++) {
          let reports = xmlDocument.documentElement.getElementsByTagName('Report')[i];
          for (let j = 0; j < reports.getElementsByTagName('SalesPers').length; j++) {
            let report = reports.getElementsByTagName('SalesPers')[j].innerHTML;
            chart.ChartData._get_Item(i + 1, 0).Text = report;
          }
          for (let j = 0; j < reports.getElementsByTagName('SaleAmt').length; j++) {
            let report = reports.getElementsByTagName('SaleAmt')[j].innerHTML;
            chart.ChartData._get_Item(i + 1, 1).NumberValue = report;
          }
          for (let j = 0; j < reports.getElementsByTagName('ComPct').length; j++) {
            let report = reports.getElementsByTagName('ComPct')[j].innerHTML;
            chart.ChartData._get_Item(i + 1, 3).NumberValue = report;
          }
          for (let j = 0; j < reports.getElementsByTagName('ComAmt').length; j++) {
            let report = reports.getElementsByTagName('ComAmt')[j].innerHTML;
            chart.ChartData._get_Item(i + 1, 2).NumberValue = report;
          }
        }

        //Load data from datatable to chart
        chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("B1", "D1");
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "A7");
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("B2", "B7");
        chart.Series.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(0).Fill.SolidColor.KnownColor = wasmModule.KnownColors.Brown;
        chart.Series.get_Item(1).Values = chart.ChartData._get_ItemNE("C2", "C7");
        chart.Series.get_Item(1).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(1).Fill.SolidColor.KnownColor = wasmModule.KnownColors.Green;
        chart.Series.get_Item(2).Values = chart.ChartData._get_ItemNE("D2", "D7");
        chart.Series.get_Item(2).Fill.FillType = wasmModule.FillFormatType.Solid;
        chart.Series.get_Item(2).Fill.SolidColor.KnownColor = wasmModule.KnownColors.Orange;

        //Set the 3D rotation
        chart.RotationThreeD.XDegree = 10;
        chart.RotationThreeD.YDegree = 10;

        // Define the output file name
        const outputFileName = "CreateCylinder3DClusteredChart_out.pptx";

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
