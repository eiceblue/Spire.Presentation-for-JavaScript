<template>
  <span>Click the following button to extract text in PPT document.</span>
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
        let inputFileName = "ExtractText.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Initialize an empty array to store extracted text
        let sb = [];

        // Foreach the slide and extract text
        for (let i = 0;i < ppt.Slides.Count;i++){
          // Get the current slide
          let slide = ppt.Slides.get_Item(i);
          for (let j = 0;j < slide.Shapes.Count;j++){
            // Get the current shape on the slide
            let shape = slide.Shapes.get_Item(j);
            if(shape instanceof wasmModule.IAutoShape){
              // Get the paragraphs in the text frame
              let tp = shape.TextFrame.Paragraphs;
              for (let k = 0;k < tp.Count;k++){
                // Add the text of each paragraph to the array
                sb.push(tp.get_Item(k).Text + "\n");
              }
            }
          }
        }
        // Join all extracted text into a single string
        let str = sb.join("");

        const outputFileName = "ExtractText.txt";

        // Save the content to the specified path
        wasmModule.FS.writeFile(outputFileName, str);

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
