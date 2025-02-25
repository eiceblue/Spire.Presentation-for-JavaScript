<template>
  <span>The example shows how to add exit animation effect for shape in a PPT document. </span>
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
        const ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        let rect =  wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        slide.Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle, fileName:ImageFileName, rectangle:rect});
        slide.Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color =  wasmModule.Color.get_FloralWhite();

        //Add a shape to the slide
        let starShape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.FivePointedStar,rectangle: wasmModule.RectangleF.FromLTRB(250, 100, 450, 300)});
        starShape.Fill.FillType = wasmModule.FillFormatType.Solid;
        starShape.Fill.SolidColor.KnownColor = wasmModule.KnownColors.LightBlue;

        //Add random bars effect to the shape
        let effect = slide.Timeline.MainSequence.AddEffect(starShape, wasmModule.AnimationEffectType.RandomBars);

        //Change effect type from entrance to exit
        effect.PresetClassType = wasmModule.TimeNodePresetClassType.Exit;

        // Define the output file name
        const outputFileName = "AddExitAnimationForShape_out.pptx";

        // Save the document to the specified path
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013});

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
