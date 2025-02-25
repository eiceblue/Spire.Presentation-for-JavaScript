<template>
  <span>The example demonstrates how to set 3D effect for shapes in a PPT document. </span>
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

        const ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Set background image
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Add shape1 and fill it with color
        let shape1 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.RoundCornerRectangle, rectangle:wasmModule.RectangleF.FromLTRB(150, 150, 300, 300)});
        shape1.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape1.Fill.SolidColor.KnownColor = wasmModule.KnownColors.SkyBlue;

        //Initialize a new instance of the 3-D class for shape1 and set its properties
        let effect1 = shape1.ThreeD.ShapeThreeD;
        effect1.PresetMaterial = wasmModule.PresetMaterialType.Powder;
        effect1.TopBevel.PresetType = wasmModule.BevelPresetType.ArtDeco;
        effect1.TopBevel.Height = 4;
        effect1.TopBevel.Width = 12;
        effect1.BevelColorMode = wasmModule.BevelColorType.Contour;
        effect1.ContourColor.KnownColor = wasmModule.KnownColors.LightBlue;
        effect1.ContourWidth = 3.5;

        //Add shape2 and fill it with color
        let shape2 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Pentagon,rectangle: wasmModule.RectangleF.FromLTRB(400, 150, 550, 300)});
        shape2.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape2.Fill.SolidColor.KnownColor = wasmModule.KnownColors.LightGreen;

        //Initialize a new instance of the 3-D class for shape2 and set its properties
        let effect2 = shape2.ThreeD.ShapeThreeD;
        effect2.PresetMaterial = wasmModule.PresetMaterialType.SoftEdge;
        effect2.TopBevel.PresetType = wasmModule.BevelPresetType.SoftRound;
        effect2.TopBevel.Height = 12;
        effect2.TopBevel.Width = 12;
        effect2.BevelColorMode = wasmModule.BevelColorType.Contour;
        effect2.ContourColor.KnownColor = wasmModule.KnownColors.LawnGreen;
        effect2.ContourWidth = 5;

        // Define the output file name
        const outputFileName = "Set3DEffectForShape_out.pptx";

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
