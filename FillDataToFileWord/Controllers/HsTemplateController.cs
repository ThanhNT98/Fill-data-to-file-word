using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.Http;
using Word = Microsoft.Office.Interop.Word;

namespace FillDataToFileWord.Controllers
{
	public class FileUtils
	{
		public static string GetContentType(string path)
		{
			var types = GetMimeTypes();
			var ext = Path.GetExtension(path).ToLowerInvariant();
			return types[ext];
		}

		public static Dictionary<string, string> GetMimeTypes()
		{
			return new Dictionary<string, string>
			{
				{".aac", "audio/aac"},
				{".abw", "application/x-abiword"},
				{".arc", "application/x-freearc"},
				{".avi", "video/x-msvideo"},
				{".azw", "application/vnd.amazon.ebook"},
				{".bin", "application/octet-stream"},
				{".bmp", "image/bmp"},
				{".bz", "application/x-bzip"},
				{".bz2", "application/x-bzip2"},
				{".csh", "application/x-csh"},
				{".css", "text/css"},
				{".csv", "text/csv"},
				{".doc", "application/msword"},
				{".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
				{".eot", "application/vnd.ms-fontobject"},
				{".epub", "application/epub+zip"},
				{".gif", "image/gif"},
				{".html", "text/html"},
				{".htm", "text/html"},
				{".ico", "image/vnd.microsoft.icon"},
				{".ics", "text/calendar"},
				{".jar", "application/java-archive"},
				{".jpeg", "image/jpeg"},
				{".jpg", "image/jpeg"},
				{".js", "text/javascript"},
				{".json", "application/json"},
				{".mid", "audio/x-midi"},
				{".midi", "audio/x-midi"},
				{".mjs", "text/javascript"},
				{".mp3", "audio/mpeg"},
				{".mpeg", "video/mpeg"},
				{".mpkg", "application/vnd.apple.installer+xml"},
				{".odp", "application/vnd.oasis.opendocument.presentation"},
				{".ods", "application/vnd.oasis.opendocument.spreadsheet"},
				{".odt", "application/vnd.oasis.opendocument.text"},
				{".oga", "audio/ogg"},
				{".ogv", "video/ogg"},
				{".ogx", "application/ogg"},
				{".otf", "font/otf"},
				{".png", "image/png"},
				{".pdf", "application/pdf"},
				{".ppt", "application/vnd.ms-powerpoint"},
				{".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"},
				{".rar", "application/x-rar-compressed"},
				{".rtf", "application/rtf"},
				{".sh", "application/x-sh"},
				{".svg", "image/svg+xml"},
				{".swf", "application/x-shockwave-flash"},
				{".tar", "application/x-tar"},
				{".tif","image/tiff"},
				{" tiff", "image/tiff"},
				{".ttf", "font/ttf"},
				{".txt", "text/plain"},
				{".vsd", "application/vnd.visio"},
				{".wav", "audio/wav"},
				{".weba", "audio/webm"},
				{".webm", "video/webm"},
				{".webp", "image/webp"},
				{".woff", "font/woff"},
				{".woff2", "font/woff2"},
				{".xhtml", "application/xhtml+xml"},
				{".xls", "application/vnd.ms-excel"},
				{".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
				{".xml", "application/xml "},
				{".zip", "application/zip"},
				{".3gp", "video/3gpp"},
				{".3g2", "video/3gpp2"},
				{".7z", "application/x-7z-compressed"}
			};
		}
	}
	public class HsTemplateController : ApiController
    {
		[HttpPost()]
		public IHttpActionResult UploadTemplate()
		{
			Word.Application wordApp = null;
			Word.Document wordDoc = null;

			try
			{
				var httpRequest = HttpContext.Current.Request;
				var path = "";
				foreach (string file in httpRequest.Files)
				{
					var fileNames = httpRequest.Files[file].FileName;
					if (fileNames.Contains("=?utf-8?B?"))
					{
						fileNames = fileNames.Replace("=?utf-8?B?", "").Replace("?=", "");
						var bytes = Convert.FromBase64String(fileNames);
						fileNames = Encoding.UTF8.GetString(bytes);
					}

					path = HttpContext.Current.Server.MapPath("~/Template/" + fileNames);
					httpRequest.Files[file].SaveAs(path);
				}

				ICollection<string> fields = new List<string>();

				var missing = Missing.Value;
				wordApp = new Word.Application();
				wordDoc = wordApp.Documents.Add(path, missing, missing, missing);

				foreach (Word.ContentControl contentControl in wordDoc.ContentControls)
				{
					try
					{
						fields.Add(contentControl.Tag.Trim());
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
								fields.Add(contentControl.Tag.Trim());
							}
							catch (Exception) { }
						}
					}
					catch (Exception) { }
				}

				return Ok(fields);
			}
			catch (Exception e)
			{
				return BadRequest(e.Message);
			}
			finally
			{
				if (wordDoc != null)
				{
					wordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges,
					Word.WdOriginalFormat.wdOriginalDocumentFormat,
					false);
				}
				if (wordApp != null)
				{
					wordApp.Quit(false);
				}
			}
		}

	}
}