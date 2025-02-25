<template>
  <span>Click the following button to hide audio during show.</span>
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
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Hide Audio during show
        for (let i = 0; i < slide.Shapes.Count; i++) {
            let shape = slide.Shapes.get_Item(i);
            if(shape instanceof wasmModule.IAudio){
                let audio = shape;
                audio.HideAtShowing = true;
            }
        }

        // Define the output file name
        const outputFileName = "HideAudioDuringShow_out.pptx";

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
