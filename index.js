const express = require("express");
const htmlDocs = require("html-docx-js");
const fs = require("fs");
const path = require("path");

console.log("dhruv");
const app = express();
let htmlContent = `<p style="text-align:center">Hi<br />[[u-i-d]]<br /><strong>Lorem Ipsum&nbsp;is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry&#39;s standard dummy</strong> text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.</p><table align="center" border="1" cellpadding="1" cellspacing="1" style="width:500px" summary="Summary"><caption>Caption</caption><thead><tr>
<th scope="col">Header1</th>
<th scope="col">Header2</th>
<th scope="col">Header3</th>
<th scope="col">Header4</th></tr></thead><tbody><tr><td>data1</td><td>data2</td><td>data3</td><td>data4</td></tr><tr><td>data5</td><td>data6</td>
<td>data7</td><td>data8</td></tr></tbody></table><p>&nbsp;</p>
  `;
let tableContent = `
<p>&nbsp;</p>
<table align="center" border="1" cellpadding="1" cellspacing="1" style="width:500px" summary="Summary">
	<thead>
		<tr>
			<th scope="col">Header1</th>
			<th scope="col">Header2</th>
			<th scope="col">Header3</th>
			<th scope="col">Header4</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>data1</td>
			<td>data2</td>
			<td>data3</td>
			<td>data4</td>
		</tr>
		<tr>
			<td>data5</td>
			<td>data6</td>
			<td>data7</td>
			<td>data8</td>
		</tr>
	</tbody>
</table>
<p>&nbsp;</p>

`;
app.get("/", async (req, res, next) => {
  try {
    htmlContent = htmlContent.replace("[[u-i-d]]", " ABCD");
    const converted = htmlDocs.asBlob(htmlContent);
    // const saveFile = await fs.writeFileSync("test.docx", converted);
    const arrayBuffer = await converted.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    // Write the Buffer to a DOCX file
    const folderPath = path.join(__dirname, "/dhruv");
    fs.mkdirSync(folderPath, { recursive: true });
    const docxFilePath = path.join(folderPath, "output.docx");
    const saveFile = fs.writeFileSync(docxFilePath, buffer);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", "attachment; filename=output.docx");

    // Read the file and send it as the response
    const fileStream = fs.createReadStream("dhruv/output.docx");
    fileStream.pipe(res);
    // return res.status(200).json({ saveFile });
  } catch (error) {
    console.log("Error", error);
  }
});

app.listen(3000, () => {
  console.log("App running on 3000");
});
