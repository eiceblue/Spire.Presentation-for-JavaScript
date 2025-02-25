<template>
  <span>Click the following button to retain notes while converting PPT to SVG.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";
import JSZip from "jszip";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName = "Template_Ppt_5.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the file from disk.
        ppt.LoadFromFile(inputFileName);

        //Retain the notes while converting PowerPoint file to svg file.
        ppt.IsNoteRetained = true;
        let outFileName = "";
        for (let i = 0; i < ppt.Slides.Count; i++) {
            //Convert presentation slides to svg file.
            let bytes = ppt.Slides.get_Item(i).SaveToSVG();
            outFileName = `output_${i}.svg`;          
            bytes.Save(outFileName);

        }

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

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
