<template>
  <span>The below example shows how to remove images of PPT</span>
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
        const inputFileName = "RemoveImages.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);
        
        // Get the first slide
        let slide = ppt.Slides.get_Item(0);

        for (let i = slide.Shapes.Count - 1; i >= 0; i--) {
          //It is the SlidePicture object
          if (slide.Shapes.get_Item(i) instanceof wasmModule.SlidePicture) {
            slide.Shapes.RemoveAt(i);
          }
        }

        // Define the output file name
        const outputFileName = "RemoveImages.pptx";

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
