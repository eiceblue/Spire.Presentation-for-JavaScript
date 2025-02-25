<template>
  <span>The example demonstrates how to split a PPT document into individual slides.</span>
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

        const inputFileName = "InputTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load file
        ppt.LoadFromFile(inputFileName);

        for (let i = 0; i < ppt.Slides.Count; i++) {

          //Initialize another instance of Presentation, and remove the blank slide
          let newppt = wasmModule.Presentation.Create();
          newppt.Slides.RemoveAt(0);

          //Append the specified slide from old presentation to the new one
          newppt.Slides.Append({slide:ppt.Slides.get_Item(i)});

          //Save the document
          let outputFileName = `SplitPPT-${i}.pptx`;
          newppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });
          newppt.Dispose();

          // Read the saved file and convert to a Blob object
          const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
          const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

          // Download the file
          downloadName.value = outputFileName;
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
