<template>
  <span>Click the following button to set alignment for text in table.</span>
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

        let inputFileName = "SetAlignmentInTable.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
            let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
            if (shape instanceof wasmModule.ITable) {
                let  table = shape;
                //Horizontal Alignment
                //Set the horizontal alignment for the cells in first column
                table.get_Item(0,1).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Left;
                table.get_Item(0,2).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Center;
                table.get_Item(0,3).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Right;
                table.get_Item(0,4).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Justify;

                //Vertical Alignment
                //Set the vertical alignment for the cells in second column
                table.get_Item(1,1).TextAnchorType = wasmModule.TextAnchorType.Top;
                table.get_Item(1,2).TextAnchorType = wasmModule.TextAnchorType.Center;
                table.get_Item(1,3).TextAnchorType = wasmModule.TextAnchorType.Bottom;
                table.get_Item(1,4).TextAnchorType = wasmModule.TextAnchorType.None;

                //Both orientaions
                //Set the both horizontal and vertical alignment for the cells in the third column
                table.get_Item(2,1).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Left;
                table.get_Item(2,1).TextAnchorType = wasmModule.TextAnchorType.Top;

                table.get_Item(2,2).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Right;
                table.get_Item(2,2).TextAnchorType = wasmModule.TextAnchorType.Center;

                table.get_Item(2,3).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Justify;
                table.get_Item(2,3).TextAnchorType = wasmModule.TextAnchorType.Bottom;

                table.get_Item(2,4).TextFrame.Paragraphs.get_Item(0).Alignment = wasmModule.TextAlignmentType.Center;
                table.get_Item(2,4).TextAnchorType = wasmModule.TextAnchorType.Top;
            }
        }

        // Define the output file name
        const outputFileName = "SetAlignmentInTable_out.pptx";

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
