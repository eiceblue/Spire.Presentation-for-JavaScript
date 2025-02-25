<template>
  <span>The example demonstrates how to get builtin properties in a PPT document. </span>
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

        const inputFileName = "GetProperties.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT document from disk
        ppt.LoadFromFile(inputFileName);

        //Get the builtin properties
        let application = ppt.DocumentProperty.Application;
        let author = ppt.DocumentProperty.Author;
        let company = ppt.DocumentProperty.Company;
        let keywords = ppt.DocumentProperty.Keywords;
        let comments = ppt.DocumentProperty.Comments;
        let category = ppt.DocumentProperty.Category;
        let title = ppt.DocumentProperty.Title;
        let subject = ppt.DocumentProperty.Subject;

        //Create StringBuilder to save
        let content = [];
        content.push("DocumentProperty.Application: " + application);
        content.push("DocumentProperty.Author: " + author);
        content.push("DocumentProperty.Company " + company);
        content.push("DocumentProperty.Keywords: " + keywords);
        content.push("DocumentProperty.Comments: " + comments);
        content.push("DocumentProperty.Category: " + category);
        content.push("DocumentProperty.Title: " + title);
        content.push("DocumentProperty.Subject: " + subject);

        // Define the output file name
        const outputFileName = "GetBuiltinProperties_out.txt";
        wasmModule.FS.writeFile(outputFileName, content.join(""));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],  {type: "text/plain"});

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
