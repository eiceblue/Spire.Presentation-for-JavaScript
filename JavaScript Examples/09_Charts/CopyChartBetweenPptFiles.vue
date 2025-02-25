<template>
  <span>Click the following button to copy chart between PowerPoint documents. </span>
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
        let inputFileName1 = 'Template_Ppt_2.pptx';
        await wasmModule.FetchFileToVFS(inputFileName1, '', `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = 'Template_Ppt_1.pptx';
        await wasmModule.FetchFileToVFS(inputFileName2, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt1 = wasmModule.Presentation.Create();

        //Load the file 
        ppt1.LoadFromFile(inputFileName1);

        //Get the chart that is going to be copied.
        let chart = ppt1.Slides.get_Item(0).Shapes.get_Item(0);

        //Load the second PowerPoint document.
        let ppt2 = wasmModule.Presentation.Create();
        ppt2.LoadFromFile(inputFileName2);

        //Copy chart from the first document to the second document.
        ppt2.Slides.Append();
        ppt2.Slides.get_Item(1).Shapes.CreateChart(chart, wasmModule.RectangleF.FromLTRB(100, 100, 600, 400), -1);
        
        // Define the output file name
        const outputFileName = "CopyChartBetweenPptFiles_out.pptx";

        // Save the document to the specified path
        ppt2.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

        // Clean up resources
        ppt1.Dispose();
        ppt2.Dispose();

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
