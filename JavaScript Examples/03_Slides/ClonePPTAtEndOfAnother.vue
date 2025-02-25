<template>
  <span>Click the following button to clone all slides of one PPT to the end of another PPT.</span>
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
        let inputFileName1 = "ChangeSlidePosition.pptx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);

        let inputFileName2 = "PPTSample_N.pptx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        // Load PPT document from the specified input file
        let ppt1 = wasmModule.Presentation.Create();
        ppt1.LoadFromFile(inputFileName1);

        // Load PPT document from the specified input file
        let ppt2 = wasmModule.Presentation.Create();
        ppt2.LoadFromFile(inputFileName2);

        // Loop through all slides of source document
        for (let i = 0; i < ppt1.Slides.Count; i++) {
          // Append the slide at the end of destination document
          let slide = ppt1.Slides.get_Item(i);
          ppt2.Slides.Append({ slide: slide });
        }

        const outputFileName = "ClonePPTAtEndOfAnother.pptx";

        // Save to file
        ppt2.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

        // Clean up resources
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
