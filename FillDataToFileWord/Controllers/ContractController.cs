using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

using Word = Microsoft.Office.Interop.Word;

namespace FillDataToFileWord.Controllers
{
    public class ContractController : ApiController
    {
		[HttpPost()]
		public async Task<HttpResponseMessage> GetContractAsync()
		{
			Word.Application wordApp = null;
			Word.Document wordDoc = null;

			try
			{
				var requestContent = Request.Content.ReadAsStringAsync().Result;
				var data = JsonConvert.DeserializeObject<Dictionary<string, string>>(requestContent);

				var missing = Missing.Value;
				wordApp = new Word.Application();
				if (!data.TryGetValue("FileName", out string fileName))
				{
					return Request.CreateResponse(System.Net.HttpStatusCode.NotFound);
				}

				string path = HttpContext.Current.Server.MapPath("~/Template/" + fileName);
				wordDoc = wordApp.Documents.Add(path, missing, missing, missing);

				foreach (Word.ContentControl contentControl in wordDoc.ContentControls)
				{
					try
					{
						string tag = contentControl.Tag.Trim();
						if (tag.Contains("ChkBox"))
							contentControl.Checked = data.ContainsKey(tag);
						else
						{
							contentControl.Range.Select();
							data.TryGetValue(tag, out string value);
							wordApp.Selection.TypeText(value != null ? value : "...............");
						}
					}
					catch (Exception) { }
				}

				foreach (Word.Shape shape in wordDoc.Shapes)
				{
					try
					{
						foreach (Word.ContentControl contentControl in shape.TextFrame.TextRange.ContentControls)
						{
							try
							{
								string tag = contentControl.Tag.Trim();
								if (tag.Contains("ChkBox"))
									contentControl.Checked = data.ContainsKey(tag);
								else
								{
									contentControl.Range.Select();
									data.TryGetValue(tag, out string value);
									wordApp.Selection.TypeText(value != null ? value : "...............");
								}
							}
							catch (Exception) { }
						}
					}
					catch (Exception) { }
				}

				string savePath = HttpContext.Current.Server.MapPath("~/Document/" + fileName);
				wordDoc.SaveAs2(savePath);
				wordDoc.Close();
				wordApp.Quit();

				var memory = new MemoryStream();
				using (var stream = new FileStream(savePath, FileMode.Open))
				{
					await stream.CopyToAsync(memory);
				}
				memory.Position = 0;

				var result = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
				{
					Content = new ByteArrayContent(memory.ToArray())
				};
				result.Content.Headers.ContentDisposition =
					 new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
					 {
						 FileName = fileName
					 };
				result.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue
					("application/octet-stream");
				return result;

			}
			catch (Exception e)
			{
				return Request.CreateResponse(System.Net.HttpStatusCode.BadRequest);
			}
			finally
			{
				try
				{
					if (wordDoc != null)
					{
						wordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges,
						Word.WdOriginalFormat.wdOriginalDocumentFormat,
						false);
					}
				}
				catch (Exception) { }
				try
				{
					if (wordApp != null)
					{
						wordApp.Quit(false);
					}
				}
				catch (Exception) { }
			}
		}
	}
}