<template>
  <span>Click the following button to create multiple paragraphs.</span>
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
        let inputFileName = "Template_Az.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Access the first slide
        let slide = ppt.Slides.get_Item(0);

        // Add an AutoShape of rectangle type
        let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 250, 150, (500 + ppt.SlideSize.Size.Width / 2 - 250), 300);
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: rec });

        // Access TextFrame of the AutoShape
        let tf = shape.TextFrame;

        // Create Paragraphs and TextRanges with different text formats
        let para0 = tf.Paragraphs.get_Item(0);
        let textRange1 = wasmModule.TextRange.Create("");
        let textRange2 = wasmModule.TextRange.Create("");
        para0.TextRanges.Append(textRange1);
        para0.TextRanges.Append(textRange2);

        let para1 = wasmModule.TextParagraph.Create();
        tf.Paragraphs._Append(para1);
        let textRange11 = wasmModule.TextRange.Create("");
        let textRange12 = wasmModule.TextRange.Create("");
        let textRange13 = wasmModule.TextRange.Create("");
        para1.TextRanges.Append(textRange11);
        para1.TextRanges.Append(textRange12);
        para1.TextRanges.Append(textRange13);

        let para2 = wasmModule.TextParagraph.Create();
        tf.Paragraphs._Append(para2);
        let textRange21 = wasmModule.TextRange.Create("");
        let textRange22 = wasmModule.TextRange.Create("");
        let textRange23 = wasmModule.TextRange.Create("");
        para2.TextRanges.Append(textRange21);
        para2.TextRanges.Append(textRange22);
        para2.TextRanges.Append(textRange23);

        // Iterate through the first three paragraphs
        for (let i = 0; i < 3; i++) {
          // Iterate through the first three text ranges in each paragraph
          for (let j = 0; j < 3; j++) {
            // Set the text for each text range
            tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Text = "TextRange " + j.toString();
            // Apply formatting based on the index of the text range
            if (j == 0) {
              // Format for the first text range
              tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Fill.FillType = wasmModule.FillFormatType.Solid;
              tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
              tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Format.IsBold = wasmModule.TriState.True;
              tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).FontHeight = 15;
            }
            else if (j == 1) {
              // Format for the second text range
              tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Fill.FillType = wasmModule.FillFormatType.Solid;
              tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Fill.SolidColor.Color = wasmModule.Color.get_Blue();
              tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).Format.IsItalic = wasmModule.TriState.True;
              tf.Paragraphs.get_Item(i).TextRanges.get_Item(j).FontHeight = 18;
            }
          }
        }

        const outputFileName = "MultipleParagraphs.pptx";

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
