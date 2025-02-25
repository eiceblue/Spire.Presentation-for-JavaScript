<template>
  <span>Click the following button to merge selected slides to a single PPT document.</span>
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

        // Load the sample files into the virtual file system (VFS)
        let inputFileName1 = "InputTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);

        let inputFileName2 = "TextTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create an instance of presentation document
        let ppt = wasmModule.Presentation.Create();

        // Remove the first slide
        ppt.Slides.RemoveAt(0);

        // Load PPT document from the specified input file
        let ppt1 = wasmModule.Presentation.Create();
        ppt1.LoadFromFile(inputFileName1);

        // Load PPT document from the specified input file
        let ppt2 = wasmModule.Presentation.Create();
        ppt2.LoadFromFile(inputFileName2);

        // Append all slides in ppt1 to ppt
        for (let i = 0; i < ppt1.Slides.Count; i++) {
          ppt.Slides.Append({ slide: ppt1.Slides.get_Item(i) });
        }

        // Append the second slide in ppt2 to ppt
        ppt.Slides.Append({ slide: ppt2.Slides.get_Item(1) });

        const outputFileName = "MergeSelectedSlides.pptx";

        // Save to file
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
