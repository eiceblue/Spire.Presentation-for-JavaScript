<template>
  <span>Click the following button to mix font styles within a single text range.</span>
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
        let inputFileName = "FontStyle.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Get the second shape of the first slide
        let shape = ppt.Slides.get_Item(0).Shapes.get_Item(1);

        // Get the text from the shape
        let originalText = shape.TextFrame.Text;

        // Split the string by specified words and return substrings to a string array
        let keywords = ["bold", "red", "underlined", "bigger font size"];
        let regex = new RegExp(keywords.map(keywords => keywords.replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&')).join('|'),'g');
        let splitArray = originalText.split(regex).filter(Boolean);

        // Remove the paragraph from TextRange
        let tp = shape.TextFrame.TextRange.Paragraph;
        tp.TextRanges.Clear();

        // Append normal text that is in front of 'bold' to the paragraph
        let tr = wasmModule.TextRange.Create(splitArray[0]);
        tp.TextRanges.Append(tr);

        // Set font style of the text 'bold' as bold
        tr = wasmModule.TextRange.Create("bold");
        tr.IsBold = wasmModule.TriState.True;
        tp.TextRanges.Append(tr);

        // Append normal text that is in front of 'red' to the paragraph
        tr = wasmModule.TextRange.Create(splitArray[1]);
        tp.TextRanges.Append(tr);

        // Set the color of the text 'red' as red
        tr = wasmModule.TextRange.Create("red");
        tr.Fill.FillType = wasmModule.FillFormatType.Solid;
        tr.Format.Fill.SolidColor.Color = wasmModule.Color.get_Red();
        tp.TextRanges.Append(tr);

        // Append normal text that is in front of 'underlined' to the paragraph
        tr = wasmModule.TextRange.Create(splitArray[2]);
        tp.TextRanges.Append(tr);

        // Underline the text 'undelined'
        tr = wasmModule.TextRange.Create("underlined");
        tr.TextUnderlineType = wasmModule.TextUnderlineType.Single;
        tp.TextRanges.Append(tr);

        // Append normal text that is in front of 'bigger font size' to the paragraph
        tr = wasmModule.TextRange.Create(splitArray[3]);
        tp.TextRanges.Append(tr);
        
        // Set a large font for the text 'bigger font size'
        tr = wasmModule.TextRange.Create("bigger font size");
        tr.FontHeight = 35;
        tp.TextRanges.Append(tr);

        // Append other normal text
        tr = wasmModule.TextRange.Create(splitArray[4]);
        tp.TextRanges.Append(tr);

        const outputFileName = "MixFontStyles.pptx";

        // Save to file
        ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

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
