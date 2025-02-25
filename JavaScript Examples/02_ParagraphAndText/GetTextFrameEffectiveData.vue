<template>
  <span>Click the following button to get TextFrame effective data.</span>
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
        let inputFileName = "Template_Az1.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Get a first shape from the slide
        let shape = slide.Shapes.get_Item(0);

        // Get the TextFrame from shape
        let textFrameFormat = shape.TextFrame;

        // Initialize an empty array to store extracted text
        let str = [];

        // Add the anchoring type of the text frame to the string
        str.push("Anchoring type: " + textFrameFormat.AnchoringType + "\n");

        // Add the autofit type of the text frame to the string
        str.push("Autofit type: " + textFrameFormat.AutofitType + "\n");

        // Add the vertical text type of the text frame to the string
        str.push("Text vertical type: " + textFrameFormat.VerticalTextType + "\n");

        // Add the margins of the text frame to the string
        str.push("Margins" + "\n");
        str.push("   Left: " + textFrameFormat.MarginLeft + "\n");
        str.push("   Top: " + textFrameFormat.MarginTop + "\n");
        str.push("   Right: " + textFrameFormat.MarginRight + "\n");
        str.push("   Bottom: " + textFrameFormat.MarginBottom + "\n");

        // Join all extracted text into a single string
        let content = str.join("");

        const outputFileName = "GetTextFrameEffectiveData.txt";

        // Save the content to the specified path
        wasmModule.FS.writeFile(outputFileName, content);

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