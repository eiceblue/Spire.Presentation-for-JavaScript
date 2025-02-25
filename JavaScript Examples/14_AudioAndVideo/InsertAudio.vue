<template>
  <span>Click the following button to insert audio into a PPT document.</span>
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

        let inputFileName = "InsertAudio.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        const audioFileName = "Music.wav";
        await wasmModule.FetchFileToVFS(audioFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the document from disk
        ppt.LoadFromFile(inputFileName);

        //Add title
        let rec_title = wasmModule.RectangleF.FromLTRB(50, 240, 210, 290);
        let shape_title = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rec_title});
        shape_title.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();

        shape_title.Fill.FillType = wasmModule.FillFormatType.None;
        let para_title = wasmModule.TextParagraph.Create();
        para_title.Text = "Audio:";
        para_title.Alignment = wasmModule.TextAlignmentType.Center;
        para_title.TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Myriad Pro Light");
        para_title.TextRanges.get_Item(0).FontHeight = 32;
        para_title.TextRanges.get_Item(0).IsBold = wasmModule.TriState.True;
        para_title.TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        para_title.TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.FromArgb(68, 68, 68);
        shape_title.TextFrame.Paragraphs._Append(para_title);

        //Insert audio into the document
        let audioRect = wasmModule.RectangleF.FromLTRB(220, 240, 300, 320);
        ppt.Slides.get_Item(0).Shapes.AppendAudioMedia({filePath:audioFileName, rectangle:audioRect});

        // Define the output file name
        const outputFileName = "InsertAudio_out.pptx";

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
