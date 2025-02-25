<template>
  <span>The example demonstrates how to set properties for a template. </span>
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
  
        // Define the output file name
        const outputFileName = "SetPropertiesForTemplate_out.odp";

        SetPropertiesForTemplate(outputFileName, wasmModule.FileFormat.ODP)

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.oasis.opendocument.presentation" });

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
      function  SetPropertiesForTemplate(fileName, fileFormat) {
        //Create a document
        let ppt = wasmModule.Presentation.Create();

        //Set the DocumentProperty
        ppt.DocumentProperty.Application = "Spire.Presentation";
        ppt.DocumentProperty.Author = "E-iceblue";
        ppt.DocumentProperty.Company = "E-iceblue Co., Ltd.";
        ppt.DocumentProperty.Keywords = "Demo File";
        ppt.DocumentProperty.Comments = "This file is used to test Spire.Presentation.";
        ppt.DocumentProperty.Category = "Demo";
        ppt.DocumentProperty.Title = "This is a demo file.";
        ppt.DocumentProperty.Subject = "Test";

        //Save to template file
        ppt.SaveToFile({file:fileName,fileFormat:fileFormat});
        ppt.Dispose();
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
