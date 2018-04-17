#r "netoffice\\NetOffice.dll"
#r "netoffice\\WordApi.dll"

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Word = NetOffice.WordApi;

public class Startup
{
    public async Task<object> Invoke(object input)
    {
      // Parse input args
      string file = "";
      string text = "";
      Dictionary<string, object> parameters = (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value); 
      object tmp;
      if (parameters.TryGetValue("file", out tmp))
      {
          file = (string)tmp;
      }
      if (parameters.TryGetValue("text", out tmp))
      {
          text = (string)tmp;
      }

      // Creat a new world instance
      Word.Application app = new Word.Application();
      // Make the window visible
      app.Visible = true;
      // Open the file
      Word.Document doc = app.Documents.Open(file);
      // Read the first section
      string outputText = doc.Sections.FirstOrDefault().Range.Text;
      // Write the first section
      doc.Sections.FirstOrDefault().Range.Text = text.ToString();
      // Save and close the document
      doc.Save();
      doc.Close();
      // Quit word app
      app.Quit();
      // Release ComObject
      app.Dispose();

      // return the old text       
      return outputText;
    }
}