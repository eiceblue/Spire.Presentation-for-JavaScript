<template>
  <span>Click the following button to get the information of comments in a PPT.</span>
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

        let inputFileName = "Comments.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        let str = [];

        //Load document from disk
        ppt.LoadFromFile(inputFileName);

        //Loop through comments
        for (let i = 0; i < ppt.CommentAuthors.Count; i++) {
            let commentAuthor = ppt.CommentAuthors.get_Item(i);
            for (let j = 0; j < commentAuthor.CommentsList.Count; j++) {
                let comment = commentAuthor.CommentsList.get_Item(j);
                //Get comment information
                let commentText = comment.Text;
                let authorName = comment.AuthorName;
                let time = comment.DateTime;
                str.push("Comment text : " + comment.Text + "\n" + "Comment author : " + comment.AuthorName + "\n" + "Posted on time : " + comment.DateTime._ToString());
            }
        }

        // Define the output file name
        const outputFileName = "GetSlideComments_out.txt";
        FS.writeFile(outputFileName, str.join("\r\n"));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "text/plain" });

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
