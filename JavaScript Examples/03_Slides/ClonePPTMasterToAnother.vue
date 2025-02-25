<template>
  <span>Click the following button to clone the master from one PPT to another PPT.</span>
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
        let inputFileName1 = "CloneMaster1.pptx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);

        let inputFileName2 = "CloneMaster2.pptx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        // Load source document from the specified input file
        let ppt1 = wasmModule.Presentation.Create();
        ppt1.LoadFromFile(inputFileName1);

        // Load destination document from the specified input file
        let ppt2 = wasmModule.Presentation.Create();
        ppt2.LoadFromFile(inputFileName2);

        // Add masters from PPT1 to PPT2
        for (let i = 0; i < ppt1.Masters.Count; i++) {
          let masterSlide = ppt1.Masters.get_Item(i);
          ppt2.Masters.AppendSlide(masterSlide);
        }

        const outputFileName = "ClonePPTMasterToAnother.pptx";

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
