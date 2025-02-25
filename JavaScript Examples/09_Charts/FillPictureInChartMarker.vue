<template>
  <span>Click the following button to fill picture in chart marker.</span>
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
        let inputFileName = 'ChartSample4.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);


        // Load the sample file into the virtual file system (VFS)
        let inputFileNameImg = "Logo.png";
        await wasmModule.FetchFileToVFS(inputFileNameImg, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        //Get chart on the first slide
        let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Load image file in ppt
        let stream = wasmModule.Stream.CreateByFile(inputFileNameImg);
        let IImage = ppt.Images.Append({ stream: stream });

        //Create a ChartDataPoint object and specify the index
        let dataPoint = wasmModule.ChartDataPoint.Create(Chart.Series.get_Item(0));
        dataPoint.Index = 0;

        //Fill picture in marker
        dataPoint.MarkerFill.Fill.FillType = wasmModule.FillFormatType.Picture;
        dataPoint.MarkerFill.Fill.PictureFill.Picture.EmbedImage = IImage;

        //Set marker size
        dataPoint.MarkerSize = 20;

        //Add the data point in series
        Chart.Series.get_Item(0).DataPoints.Add(dataPoint);

        // Define the output file name
        const outputFileName = "FillPictureInChartMarker_out.pptx";

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
