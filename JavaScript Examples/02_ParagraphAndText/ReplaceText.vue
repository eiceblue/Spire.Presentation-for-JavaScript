<template>
  <span>Click the following button to replace text in a PPT document.</span>
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
        let inputFileName = "TextTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        let tagValues = new Map();
        tagValues.set('Spire.Presentation for JavaScript','Spire.PPT');

        // Replaces specific tags in the text of shapes within a slide
        ReplaceTags(ppt.Slides.get_Item(0), tagValues);

        const outputFileName = "ReplaceText.pptx";

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

    function ReplaceTags(pSlide, TagValues) {
      for (let i = 0; i < pSlide.Shapes.Count; i++) {
        let curShape = pSlide.Shapes.get_Item(i);
        if (curShape instanceof wasmModule.IAutoShape) {
          for (let j = 0; j < curShape.TextFrame.Paragraphs.Count; j++) {
            let tp = curShape.TextFrame.Paragraphs.get_Item(j);
            for (let [key, value] of TagValues.entries()) {
              let curKey = key;
              let txt = tp.Text;
              tp.Text = findAllIndices(txt, curKey, value);
            }
          }
        }
      }
    };

    function findAllIndices(str, value, replaceValue) {
      let result = str;
      let index = 0;
      while ((index = result.indexOf(value, index)) !== -1) {
        result = result.substring(0, index) + replaceValue + result.substring(index + value.length);
        index += replaceValue.length;
      }
      return result;
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
