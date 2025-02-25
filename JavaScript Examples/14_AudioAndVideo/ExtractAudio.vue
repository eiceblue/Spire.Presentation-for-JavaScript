<template>
  <span>Click the following button to extract the audio from a slide.</span>
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

        let inputFileName = "audio.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        let outFileName = "ExtractAudio.wav";
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
            let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
            if(shape instanceof wasmModule.IAudio){
                let audio = shape;
                let AudioData = audio.Data;
                AudioData.SaveToFile(outFileName);        
            }
        }

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "audio/wav" });

        // Clean up resources
        ppt.Dispose();

        // Download the file
        downloadName.value = outFileName;
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
