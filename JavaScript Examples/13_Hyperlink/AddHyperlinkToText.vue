<template>
  <span>Click the following button to add hyperlink to text in PowerPoint document.</span>
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

        let inputFileName = "AddHyperlinkToText.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the file from disk.
        ppt.LoadFromFile(inputFileName);

        //Find the text we want to add link to it.
        let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);
        let tp = shape.TextFrame.TextRange.Paragraph;
        let temp = tp.Text;

        //Split the original text.
        let textToLink = "Spire.Presentation";
        let strSplit = temp.split(textToLink);

        //Clear all text.
        tp.TextRanges.Clear();

        //Add new text.
        let tr = wasmModule.TextRange.Create(strSplit[0]);
        tp.TextRanges.Append(tr);

        //Add the hyperlink.
        tr = wasmModule.TextRange.Create(textToLink);
        tr.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
        tp.TextRanges.Append(tr);


        // Define the output file name
        const outputFileName = "AddHyperlinkToText_out.pptx";

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
