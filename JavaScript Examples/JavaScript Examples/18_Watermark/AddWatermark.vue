<template>
  <span>Click the following button to add text watermark. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'AddWatermark.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT
        ppt.LoadFromFile(inputFileName);

        //Define a rectangle range
        let left = (ppt.SlideSize.Size.Width - 400) / 2;
        let top = (ppt.SlideSize.Size.Height - 300) / 2;
        let rect = wasmModule.RectangleF.FromLTRB(left, top, 400 + left, 300 + top);

        //Add a rectangle shape with a defined range
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({
          shapeType: wasmModule.ShapeType.Rectangle,
          rectangle: rect,
        });

        //Set the style of the shape
        shape.Fill.FillType = wasmModule.FillFormatType.None;
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
        shape.Rotation = -45;
        shape.Locking.SelectionProtection = true;
        shape.Line.FillType = wasmModule.FillFormatType.None;

        //Add text to the shape
        shape.TextFrame.Text = 'E-iceblue';
        let textRange = shape.TextFrame.TextRange;
        //Set the style of the text range
        textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        textRange.Fill.SolidColor.Color = wasmModule.Color.FromArgb({
          alpha: 120,
          baseColor: wasmModule.Color.get_HotPink(),
        });
        textRange.FontHeight = 50;
        // Define the output file name
        const outputFileName = 'Watermark.pptx';

        // Save result file
        ppt.SaveToFile({
          file: outputFileName,
          fileFormat: wasmModule.FileFormat.Pptx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });

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
