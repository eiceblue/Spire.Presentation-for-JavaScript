<template>
  <span>This sample demonstrates how to set animations in a PPT document. </span>
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

        const inputFileName = "Animations.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

         //Load the document from disk
        ppt.LoadFromFile(inputFileName);

          //Add title
        let rec_title = wasmModule.RectangleF.FromLTRB(50, 200, 250, 250);
        let shape_title = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:rec_title});
        shape_title.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();

        shape_title.Fill.FillType = wasmModule.FillFormatType.None;
        let para_title = wasmModule.TextParagraph.Create();
        para_title.Text = "Animations:";
        para_title.Alignment = wasmModule.TextAlignmentType.Center;
        para_title.TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Myriad Pro Light");
        para_title.TextRanges.get_Item(0).FontHeight = 32;
        para_title.TextRanges.get_Item(0).IsBold = wasmModule.TriState.True;
        para_title.TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        para_title.TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.FromArgb({alpha: 255,red:68, green:68,blue: 68});
        shape_title.TextFrame.Paragraphs._Append(para_title);

        //Set the animation of slide to Circle
        ppt.Slides.get_Item(0).SlideShowTransition.Type = wasmModule.TransitionType.Circle;

        //Append new shape - Triangle
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Triangle,rectangle: wasmModule.RectangleF.FromLTRB(100, 280, 180, 360)});

        //Set the color of shape
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        //Set the animation of shape
        shape.Slide.Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.Path4PointStar);

        //Append new shape - Rectangle and set animation
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(210, 280, 360, 360)});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
        shape.AppendTextFrame("Animated Shape");
        shape.Slide.Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.FadedSwivel);

        //Append new shape - Cloud and set the animation
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Cloud,rectangle: wasmModule.RectangleF.FromLTRB(390, 280, 470, 360)});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_White();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_CadetBlue();
        shape.Slide.Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.FadedZoom);

        // Define the output file name
        const outputFileName = "Animations_out.pptx";

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
