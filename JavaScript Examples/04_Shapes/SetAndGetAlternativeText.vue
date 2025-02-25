<template>
  <span>The example shows how to set and get alternative text of shapes in a PPT document.</span>
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

        const inputFileName = "ShapeTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load file
        ppt.LoadFromFile(inputFileName);

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Set the alternative text (title and description)
        slide.Shapes.get_Item(0).AlternativeTitle = "Rectangle";
        slide.Shapes.get_Item(0).AlternativeText = "This is a Rectangle";

        //Get the alternative text (title and description)
        let alternativeText = "";
        let title = slide.Shapes.get_Item(0).AlternativeTitle;
        alternativeText += "Title: " + title + "\r\n";
        let description = slide.Shapes.get_Item(0).AlternativeText;
        alternativeText += "Description: " + description;

        // Define the output file name
        const outputFileName = "SetAndGetAlternativeText_out.txt";

        wasmModule.FS.writeFile(outputFileName,alternativeText);

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
