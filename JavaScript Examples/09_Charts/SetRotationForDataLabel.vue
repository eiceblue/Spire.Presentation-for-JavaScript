<template>
  <span>The following example shows how to set rotation for data label</span>
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

        // Load the input file into the virtual file system (VFS)
        const inputFileName = "SetRotationForDataLabel.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PowerPoint document
        let ppt = wasmModule.Presentation.Create();

        // Load file from VFS
        ppt.LoadFromFile(inputFileName);

        // Get the first chart
        let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Set the rotation angle for the datalabels of first serie
        for (let i = 0; i < Chart.Series.get_Item(0).Values.Count; i++) {
          let datalabel = Chart.Series.get_Item(0).DataLabels.Add();
          datalabel.ID = i;
          datalabel.RotationAngle = 45;
        }

        // Define the output file name
        const outputFileName = "SetRotationForDataLabel.pptx";

        // Save the document 
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
