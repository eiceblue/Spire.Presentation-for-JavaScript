<template>
  <span>Click the following button to get text style effective data.</span>
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

        // Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Get a first shape from the slide
        let shape = slide.Shapes.get_Item(0);

        // Initialize an empty array to store extracted text
        let str = [];

        for (let p = 0; p < shape.TextFrame.Paragraphs.Count; p++)
        {

          let paragraph = shape.TextFrame.Paragraphs.get_Item(p);
          str.push("Text style for Paragraph " + p + " :");

          //Get the paragraph style
          str.push(" Indent: " + paragraph.Indent);
          str.push(" Alignment: " + paragraph.Alignment);
          str.push(" Font alignment: " + paragraph.FontAlignment);
          str.push(" Hanging punctuation: " + paragraph.HangingPunctuation);
          str.push(" Line spacing: " + paragraph.LineSpacing);
          str.push(" Space before: " + paragraph.SpaceBefore);
          str.push(" Space after: " + paragraph.SpaceAfter.toString());
          str.push("\r\n");
          for (let r = 0; r < paragraph.TextRanges.Count; r++)
          {
              let textRange = paragraph.TextRanges.get_Item(r);
              str.push("  Text style for Paragraph " + p + " TextRange " + r + " :");

              // Get the text range style
              str.push("    Font height: " + textRange.FontHeight);
              str.push("    Language: " + textRange.Language);
              str.push("    Font: " + textRange.LatinFont.FontName);
              str.push("");
          }
        }

        // Join all extracted text into a single string
        let content = str.join("\r\n");

        const outputFileName = "GetTextStyleEffectiveData.txt";

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