using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Security.Permissions;
using System.Collections.Generic;
using System.Security;

[assembly: AllowPartiallyTrustedCallers]

namespace GhostscriptSharp
{
	/// <summary>
	/// Wraps the Ghostscript API with a C# interface
	/// </summary>
	public class GhostscriptWrapper
	{
		#region Globals

		private static readonly string[] ARGS = new string[] {
				// Keep gs from writing information to standard output
                "-q",                     
                "-dQUIET",               
                "-dPARANOIDSAFER",       // Run this command in safe mode
                "-dBATCH",               // Keep gs from going into interactive mode
                "-dNOPAUSE",             // Do not prompt and pause for each page
                "-dNOPROMPT",            // Disable prompts for user interaction 
                //@"-sFONTPATH=" + Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts),
                "-dMaxBitmap=500000000", // Set high for better performance
				"-dNumRenderingThreads=" + Environment.ProcessorCount.ToString(), //4", // Multi-core, come-on!
                "-sDEVICE=png16m",
                "-r72",
                // Configure the output anti-aliasing, resolution, etc
                "-dAlignToPixels=0",
                "-dGridFitTT=0",
                "-dTextAlphaBits=4",
                "-dGraphicsAlphaBits=4"
		};
		#endregion


		/// <summary>
		/// Generates a thumbnail jpg for the pdf at the input path and saves it 
		/// at the output path
		/// </summary>        
        //public static void GeneratePageThumb(string _URL, string outputPath, int page, int dpix, int dpiy, int width, int height)
        //System.Drawing.Bitmap

        public static byte[] GeneratePageThumb(string _URL, int page, int dpix, int dpiy, int width, int height)
		{

            string temppath = Path.GetTempPath();
            List<string> tempfilenames = new List<string>();
            var temppdffile = string.Format(@"{0}.pdf", Guid.NewGuid());
            tempfilenames.Add(temppdffile);
            string webfilename = string.Empty;

            try
            {                
                // Open a connection
                WebPermission myWebPermission = new WebPermission(PermissionState.Unrestricted);
                myWebPermission.Demand();
                System.Net.HttpWebRequest _HttpWebRequest = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(_URL);

                _HttpWebRequest.AllowWriteStreamBuffering = true;

                // You can also specify additional header values like the user agent or the referer: (Optional)
                //_HttpWebRequest.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)"
                //_HttpWebRequest.Referer = "http://www.google.com/"

                // set timeout for 20 seconds (Optional)
                _HttpWebRequest.Timeout = 20000;

                // Request response:
                System.Net.WebResponse _WebResponse = _HttpWebRequest.GetResponse();

                // Open data stream:
                System.IO.Stream _WebStream = _WebResponse.GetResponseStream();

                using (Stream file = File.Create(temppath + temppdffile))
                {
                    CopyStream(_WebStream, file);
                    file.Close();
                }
                                                                                                                                                                                                                                                                                               
                // Cleanup
                _WebResponse.Close();
                _WebStream.Close();
                
            }
            catch (Exception _Exception)
            {
                // Error - Console.WriteLine("Exception caught in process: {0}", _Exception.ToString())
                throw new Exception("Error discover!" + _Exception.ToString());
            }

            int pagecount = getpagecount(temppath + temppdffile, _URL);
            var myUniqueFileName = string.Format(@"{0}.jpg", Guid.NewGuid());
            for (int i = 1 ; i <= pagecount ; i++)
            {
                tempfilenames.Add(temppath + (i).ToString() + myUniqueFileName);
                GeneratePageThumbs(temppath + temppdffile, temppath + "%d" + myUniqueFileName, 1, i, dpix, dpiy, width, height);
            }

            //remove original file name from list
            try
            {
                FileInfo currentFile = new FileInfo(tempfilenames[0]);
                //currentFile.Delete();
                tempfilenames.RemoveAt(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error on file: {0}\r\n   {1}", tempfilenames[0], ex.Message);
            }

            //make list to array for move to create bitmap function
            string[] images = tempfilenames.ToArray();            
            System.Drawing.Bitmap finalImage = CombineBitmap(images);
            byte[] imagebytearray = imageToByteArray(finalImage);
            
            //remove the rest of the temporary bitmaps
            foreach(string thestring in tempfilenames)
            {
                try
                {
                    //var result = MessageBox.Show(thestring);
                    FileInfo currentFile = new FileInfo(thestring);
                    currentFile.Delete();
                }
                catch (Exception ex)
                {
                    var result = MessageBox.Show("Error on file: {0}\r\n   {1}" + thestring + ex.Message);
                }
            }

            try
            {
                FileInfo currentFile = new FileInfo(temppath + temppdffile);
                currentFile.Delete();
                tempfilenames.RemoveAt(0);
            }
            catch (Exception ex)
            {
                var result = MessageBox.Show("Error on file: {0}\r\n   {1}" + temppath + temppdffile + ex.Message);
            }

            return imagebytearray;
            //return finalImage;
		}

		/// <summary>
		/// Generates a collection of thumbnail jpgs for the pdf at the input path 
		/// starting with firstPage and ending with lastPage.
		/// Put "%d" somewhere in the output path to have each of the pages numbered
		/// </summary>
    
		public static void GeneratePageThumbs(string tempfile, string outputPath, int firstPage, int lastPage, int dpix, int dpiy, int width, int height)
		{       
            if (IntPtr.Size == 4)
                API.GhostScript32.CallAPI(GetArgs(tempfile, outputPath, firstPage, lastPage, dpix, dpiy, width, height));
            else
                API.GhostScript64.CallAPI(GetArgs(tempfile, outputPath, firstPage, lastPage, dpix, dpiy, width, height));
		}

		/// <summary>
		/// Rasterises a PDF into selected format
		/// </summary>
		/// <param name="inputPath">PDF file to convert</param>
		/// <param name="outputPath">Destination file</param>
		/// <param name="settings">Conversion settings</param>
		public static void GenerateOutput(string inputPath, string outputPath, GhostscriptSettings settings)
		{
            if (IntPtr.Size == 4)
                API.GhostScript32.CallAPI(GetArgs(inputPath, outputPath, settings));
            else
                API.GhostScript64.CallAPI(GetArgs(inputPath, outputPath, settings));
		}

		/// <summary>
		/// Returns an array of arguments to be sent to the Ghostscript API
		/// </summary>
		/// <param name="inputPath">Path to the source file</param>
		/// <param name="outputPath">Path to the output file</param>
		/// <param name="firstPage">The page of the file to start on</param>
		/// <param name="lastPage">The page of the file to end on</param>
		private static string[] GetArgs(string inputPath,
			string outputPath,
			int firstPage,
			int lastPage,
			int dpix,
			int dpiy, 
            int width, 
            int height)
		{
			// To maintain backwards compatibility, this method uses previous hardcoded values.

			GhostscriptSettings s = new GhostscriptSettings();
			s.Device = Settings.GhostscriptDevices.jpeg;
			s.Page.Start = firstPage;
			s.Page.End = lastPage;
			s.Resolution = new System.Drawing.Size(dpix, dpiy);
			
			Settings.GhostscriptPageSize pageSize = new Settings.GhostscriptPageSize();
            if (width == 0 && height == 0)
            {
			    pageSize.Native = GhostscriptSharp.Settings.GhostscriptPageSizes.a7;
            }
            else
            {
                pageSize.Manual = new Size(width, height);
            }
            s.Size = pageSize;

			return GetArgs(inputPath, outputPath, s);
		}

		/// <summary>
		/// Returns an array of arguments to be sent to the Ghostscript API
		/// </summary>
		/// <param name="inputPath">Path to the source file</param>
		/// <param name="outputPath">Path to the output file</param>
		/// <param name="settings">API parameters</param>
		/// <returns>API arguments</returns>
		private static string[] GetArgs(string inputPath,
			string outputPath,
			GhostscriptSettings settings)
		{
			System.Collections.ArrayList args = new System.Collections.ArrayList(ARGS);

			if (settings.Device == Settings.GhostscriptDevices.UNDEFINED)
			{
				throw new ArgumentException("An output device must be defined for Ghostscript", "GhostscriptSettings.Device");
			}

			if (settings.Page.AllPages == false && (settings.Page.Start <= 0 && settings.Page.End < settings.Page.Start))
			{
				throw new ArgumentException("Pages to be printed must be defined.", "GhostscriptSettings.Pages");
			}

			if (settings.Resolution.IsEmpty)
			{
				throw new ArgumentException("An output resolution must be defined", "GhostscriptSettings.Resolution");
			}

			if (settings.Size.Native == Settings.GhostscriptPageSizes.UNDEFINED && settings.Size.Manual.IsEmpty)
			{
				throw new ArgumentException("Page size must be defined", "GhostscriptSettings.Size");
			}

			// Output device
			args.Add(String.Format("-sDEVICE={0}", settings.Device));

			// Pages to output
			if (settings.Page.AllPages)
			{
				args.Add("-dFirstPage=1");
			}
			else
			{
				args.Add(String.Format("-dFirstPage={0}", settings.Page.Start));
				if (settings.Page.End >= settings.Page.Start)
				{
					args.Add(String.Format("-dLastPage={0}", settings.Page.End));
				}
			}

			// Page size
			if (settings.Size.Native == Settings.GhostscriptPageSizes.UNDEFINED)
			{
				args.Add(String.Format("-dDEVICEWIDTHPOINTS={0}", settings.Size.Manual.Width));
				args.Add(String.Format("-dDEVICEHEIGHTPOINTS={0}", settings.Size.Manual.Height));
                args.Add("-dFIXEDMEDIA");
                args.Add("-dPDFFitPage");
			}
			else
			{
				args.Add(String.Format("-sPAPERSIZE={0}", settings.Size.Native.ToString()));
			}

			// Page resolution
			args.Add(String.Format("-dDEVICEXRESOLUTION={0}", settings.Resolution.Width));
			args.Add(String.Format("-dDEVICEYRESOLUTION={0}", settings.Resolution.Height));

			// Files
			args.Add(String.Format("-sOutputFile={0}", outputPath));
			args.Add(inputPath);

			return (string[])args.ToArray(typeof(string));

		}

        public static int getpagecount(string vfileName, string originalfilename)
        {
            int PgCount = 0;
            if (originalfilename.Substring(originalfilename.Length - 3).ToLowerInvariant().Equals("pdf"))
            {
                FileStream fs = new FileStream(vfileName, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs);
                string pdf = sr.ReadToEnd();
                Regex rx = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection match = rx.Matches(pdf);
                //var result = MessageBox.Show(match.Count.ToString());
                PgCount = match.Count;
                fs.Close();
             }
            try
            {
                return PgCount;
            }
            catch (Exception)
            {
                return 0;
            }
            
        }

        //private static string CreateTmpFile()
        //{
        //    string fileName = string.Empty;

        //    try
        //    {
        //        // Get the full name of the newly created Temporary file. 
        //        // Note that the GetTempFileName() method actually creates
        //        // a 0-byte file and returns the name of the created file.
        //        fileName = Path.GetTempFileName();

        //        // Craete a FileInfo object to set the file's attributes
        //        FileInfo fileInfo = new FileInfo(fileName);

        //        // Set the Attribute property of this file to Temporary. 
        //        // Although this is not completely necessary, the .NET Framework is able 
        //        // to optimize the use of Temporary files by keeping them cached in memory.
        //        fileInfo.Attributes = FileAttributes.Temporary;

        //        Console.WriteLine("TEMP file created at: " + fileName);
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Unable to create TEMP file or set its attributes: " + ex.Message);
        //    }

        //    return fileName;
        //}

        /// <summary>
        /// Copies the contents of input to output. Doesn't close either stream.
        /// </summary>
        public static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }

        public static System.Drawing.Bitmap CombineBitmap(string[] files)
        {
            //read all images into memory
            List<System.Drawing.Bitmap> images = new List<System.Drawing.Bitmap>();
            System.Drawing.Bitmap finalImage = null;

            try
            {
                int width = 0;
                int height = 0;

                foreach (string image in files)
                {
                    //create a Bitmap from the file and add it to the list
                    System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(image);

                    //update the size of the final bitmap
                    //width += bitmap.Width;
                    //height = bitmap.Height > height ? bitmap.Height : height;

                    width = bitmap.Width > width ? bitmap.Width : width;
                    height += bitmap.Height;

                    images.Add(bitmap);
                }

                //create a bitmap to hold the combined image
                finalImage = new System.Drawing.Bitmap(width, height);

                //get a graphics object from the image so we can draw on it
                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(finalImage))
                {
                    //set background color
                    g.Clear(System.Drawing.Color.Black);

                    //go through each image and draw it on the final image
                    int offset = 0;
                    foreach (System.Drawing.Bitmap image in images)
                    {
                        g.DrawImage(image,
                          new System.Drawing.Rectangle(0, offset, image.Width, image.Height));
                        offset += image.Height;
                    }
                }

                //string temppath = Path.GetTempPath();
                //var temppdffile = string.Format(@"{0}.jpg", Guid.NewGuid());

                //finalImage.Save(@temppath + temppdffile, ImageFormat.Jpeg);
                return finalImage;
            }
            catch (Exception ex)
            {
                if (finalImage != null)
                    finalImage.Dispose();

                throw ex;
            }
            finally
            {
                //clean up memory
                foreach (System.Drawing.Bitmap image in images)
                {
                    image.Dispose();
                }
            }
        }

        public static byte[] imageToByteArray(System.Drawing.Image imageIn)
        {
            MemoryStream ms = new MemoryStream();
            imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
            return ms.ToArray();
        }

	}


	/// <summary>
	/// Ghostscript settings
	/// </summary>
	public class GhostscriptSettings
	{
		private Settings.GhostscriptDevices _device;
		private Settings.GhostscriptPages _pages = new Settings.GhostscriptPages();
		private System.Drawing.Size _resolution;
		private Settings.GhostscriptPageSize _size = new Settings.GhostscriptPageSize();

		public Settings.GhostscriptDevices Device
		{
			get { return this._device; }
			set { this._device = value; }
		}

		public Settings.GhostscriptPages Page
		{
			get { return this._pages; }
			set { this._pages = value; }
		}

		public System.Drawing.Size Resolution
		{
			get { return this._resolution; }
			set { this._resolution = value; }
		}

		public Settings.GhostscriptPageSize Size
		{
			get { return this._size; }
			set { this._size = value; }
		}
	}
}

namespace GhostscriptSharp.Settings
{
	/// <summary>
	/// Which pages to output
	/// </summary>
	public class GhostscriptPages
	{
		private bool _allPages = true;
		private int _start;
		private int _end;

		/// <summary>
		/// Output all pages avaialble in document
		/// </summary>
		public bool AllPages
		{
			set
			{
				this._start = -1;
				this._end = -1;
				this._allPages = true;
			}
			get
			{
				return this._allPages;
			}
		}

		/// <summary>
		/// Start output at this page (1 for page 1)
		/// </summary>
		public int Start
		{
			set
			{
				this._allPages = false;
				this._start = value;
			}
			get
			{
				return this._start;
			}
		}

		/// <summary>
		/// Page to stop output at
		/// </summary>
		public int End
		{
			set
			{
				this._allPages = false;
				this._end = value;
			}
			get
			{
				return this._end;
			}
		}
	}

	/// <summary>
	/// Output devices for GhostScript
	/// </summary>
	public enum GhostscriptDevices
	{
		UNDEFINED,
		png16m,
		pnggray,
		png256,
		png16,
		pngmono,
		pngalpha,
		jpeg,
		jpeggray,
		tiffgray,
		tiff12nc,
		tiff24nc,
		tiff32nc,
		tiffsep,
		tiffcrle,
		tiffg3,
		tiffg32d,
		tiffg4,
		tifflzw,
		tiffpack,
		faxg3,
		faxg32d,
		faxg4,
		bmpmono,
		bmpgray,
		bmpsep1,
		bmpsep8,
		bmp16,
		bmp256,
		bmp16m,
		bmp32b,
		pcxmono,
		pcxgray,
		pcx16,
		pcx256,
		pcx24b,
		pcxcmyk,
		psdcmyk,
		psdrgb,
		pdfwrite,
		pswrite,
		epswrite,
		pxlmono,
		pxlcolor
	}

	/// <summary>
	/// Output document physical dimensions
	/// </summary>
	public class GhostscriptPageSize
	{
		private GhostscriptPageSizes _fixed;
		private System.Drawing.Size _manual;

		/// <summary>
		/// Custom document size
		/// </summary>
		public System.Drawing.Size Manual
		{
			set
			{
				this._fixed = GhostscriptPageSizes.UNDEFINED;
				this._manual = value;
			}
			get
			{
				return this._manual;
			}
		}

		/// <summary>
		/// Standard paper size
		/// </summary>
		public GhostscriptPageSizes Native
		{
			set
			{
				this._fixed = value;
				this._manual = new System.Drawing.Size(0, 0);
			}
			get
			{
				return this._fixed;
			}
		}

	}

	/// <summary>
	/// Native page sizes
	/// </summary>
	/// <remarks>
	/// Missing 11x17 as enums can't start with a number, and I can't be bothered
	/// to add in logic to handle it - if you need it, do it yourself.
	/// </remarks>
	public enum GhostscriptPageSizes
	{
		UNDEFINED,
		ledger,
		legal,
		letter,
		lettersmall,
		archE,
		archD,
		archC,
		archB,
		archA,
		a0,
		a1,
		a2,
		a3,
		a4,
		a4small,
		a5,
		a6,
		a7,
		a8,
		a9,
		a10,
		isob0,
		isob1,
		isob2,
		isob3,
		isob4,
		isob5,
		isob6,
		c0,
		c1,
		c2,
		c3,
		c4,
		c5,
		c6,
		jisb0,
		jisb1,
		jisb2,
		jisb3,
		jisb4,
		jisb5,
		jisb6,
		b0,
		b1,
		b2,
		b3,
		b4,
		b5,
		flsa,
		flse,
		halfletter
	}


}