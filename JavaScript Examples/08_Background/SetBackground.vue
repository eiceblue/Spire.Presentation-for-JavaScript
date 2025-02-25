<template>
  <span>The example demonstrates how to set background style. </span>
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

        const inputFileName = "Setbackground.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        const ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Set the background of the first slide to Gradient color
        ppt.Slides.get_Item(0).SlideBackground.Type = wasmModule.BackgroundType.Custom;
        ppt.Slides.get_Item(0).SlideBackground.Fill.FillType = wasmModule.FillFormatType.Gradient;
        ppt.Slides.get_Item(0).SlideBackground.Fill.Gradient.GradientShape = wasmModule.GradientShapeType.Linear;
        ppt.Slides.get_Item(0).SlideBackground.Fill.Gradient.GradientStyle = wasmModule.GradientStyle.FromCorner1;
        ppt.Slides.get_Item(0).SlideBackground.Fill.Gradient.GradientStops.Append({position:1,knownColor: wasmModule.KnownColors.SkyBlue});
        ppt.Slides.get_Item(0).SlideBackground.Fill.Gradient.GradientStops.Append({position:0,knownColor: wasmModule.KnownColors.White});

        //Set the background of the second slide to Solid color
        ppt.Slides.get_Item(1).SlideBackground.Type = wasmModule.BackgroundType.Custom;
        ppt.Slides.get_Item(1).SlideBackground.Fill.FillType = wasmModule.FillFormatType.Solid;
        ppt.Slides.get_Item(1).SlideBackground.Fill.SolidColor.Color = wasmModule.Color.get_SkyBlue();

        ppt.Slides.Append();

        //Set the background of the third slide to picture
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(2).SlideBackground.Fill.FillType = wasmModule.FillFormatType.Picture;
        let image = ppt.Slides.get_Item(2).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
        ppt.Slides.get_Item(2).SlideBackground.Fill.PictureFill.Picture.EmbedImage = image.PictureFill.Picture.EmbedImage;

        // Define the output file name
        const outputFileName = "SetBackground_out.pptx";

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
