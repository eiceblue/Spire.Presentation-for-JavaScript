<template>
  <span>Click the following button to reply to comment.</span>
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


        //Create ppt file
        let ppt = wasmModule.Presentation.Create();

        //Create Comment author
        let author = ppt.CommentAuthors.AddAuthor("E-iceblue", "comment");

        //Add comment
        ppt.Slides.get_Item(0).AddComment({author:author, text:"Add comment", position:wasmModule.PointF.Create(18, 25),dateTime: wasmModule.DateTime.get_Now()});
        let comment = ppt.Slides.get_Item(0).Comments[0];

        //Add reply to Comment
        if (!comment.IsReply) {
            comment.Reply(author, "Add Reply1", wasmModule.DateTime.get_Now());
            comment.Reply(author, "Add Reply2", wasmModule.DateTime.get_Now());
        }

        //delete first reply
        ppt.Slides.get_Item(0).DeleteComment({author:author, text:"Add Reply1"});


        // Define the output file name
        const outputFileName = "ReplyToComment_out.pptx";

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
