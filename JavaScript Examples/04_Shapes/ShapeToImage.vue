<template>
  <span>This sample demonstrates how to convert shapes to images in a PPT document.</span>
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

        const inputFileName = "ShapeToImage.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++){
            let images = ppt.Slides.get_Item(0).Shapes.get_Item(i).SaveAsImage();
            let outFileName = `ShapeToImage-${i}.png`;
            images.Save(outFileName);
            
            // Read the saved file and convert to a Blob object
            const modifiedFileArray = wasmModule.FS.readFile(outFileName);
            const modifiedFile = new Blob([modifiedFileArray], { type: "image/png" });

            images.Dispose();

            // Download the file
            downloadName.value = outFileName;
            downloadUrl.value = URL.createObjectURL(modifiedFile);
        }
        // Clean up resources
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
