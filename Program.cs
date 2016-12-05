using System;
using System.Collections.Generic;
using System.Linq;

namespace InDesignGenerator
{
	class Program
	{

		[CommandLine.Option('v', "verbose", Required = false, HelpText = "Template.", DefaultValue = true)]
		public bool Verbose { get; set; }

		[CommandLine.Option('t', "template-file", Required = true, HelpText = "Template.")]
		public string TemplateFile { get; set; }

		[CommandLine.Option('d', "template-dir", Required = false, HelpText = "Template.")]
		public string TemplateFileDir { get; set; }

		[CommandLine.Option('x', "working-dir", Required = false, HelpText = "Template.")]
		public string WorkingFileDir { get; set; }

		[CommandLine.Option('o', "output-file", Required = false, HelpText = "Template.")]
		public string OutputFile { get; set; }

		[CommandLine.Option('s', "content-spread-name", Required = true, HelpText = "Template.")]
		public string ContentSpreadName { get; set; }

		[CommandLine.Option('r', "variables-file", Required = false, HelpText = "Template.")]
		public string VariablesFile { get; set; }

		[CommandLine.Option('r', "photo-dir", Required = true, HelpText = "Template.")]
		public string PhotoDir { get; set; }

		public string ContentSpreadTemplate;

		[CommandLine.HelpOption]
		public string GetUsage() {
			return CommandLine.Text.HelpText.AutoBuild(this,
				(CommandLine.Text.HelpText current) => CommandLine.Text.HelpText.DefaultParsingErrorsHandler(this, current));
		}

		private Dictionary<string,string> _variables;

		public static void Main(string[] args) {
			Program program = new Program();
			if (CommandLine.Parser.Default.ParseArguments(args, program)) {
				program.Run();
			}
		}

		public void Run() {
			_debug();
			// Init
			if(this.TemplateFileDir == null) this.TemplateFileDir = this.TemplateFile + "_template";
			if(this.WorkingFileDir == null) this.WorkingFileDir = this.TemplateFile + "_tmp";
			if(this.OutputFile == null) this.OutputFile = this.TemplateFile+"_out.idml";
			if(this.VariablesFile == null) this.VariablesFile = this.TemplateFile+"_variables.txt";
			this.ContentSpreadTemplate = this.WorkingFileDir + "/" + "Spreads" + "/" + "Spread_" + this.ContentSpreadName + ".xml";
			// Process
			//UnzipTemplate();
			PrepareWorkingDir();
			PrepareVariables();
			CreateContentSpreads();
			ReplaceAllVariables();
			ZipTemplate();
			Cleanup();
		}

		public void UnzipTemplate() {
			_debug();
			string inFile = this.TemplateFile;
			string outDir = this.TemplateFileDir;
			_forceDelete(outDir);
			System.IO.Compression.ZipFile.ExtractToDirectory(inFile, outDir);
		}

		public void PrepareVariables() {
			_debug();
			// Prepare variables list
			_variables = new Dictionary<string, string>();
			foreach(var line in System.IO.File.ReadLines(this.VariablesFile)) {
				_debug(line);
				var keyval = line.Split('=');
				_variables[keyval[0]] = keyval[1];
			}
		}

		public void PrepareWorkingDir() {
			_debug();
			string inDir = this.TemplateFileDir;
			string outDir = this.WorkingFileDir;
			_forceDelete(outDir);
			_copyDir(inDir, outDir);
		}

		public void CreateContentSpreads() {
			_debug();
			// Init
			string spreadNodeTemplate = "<idPkg:Spread src=\"Spreads/[FILE]\" />";
			string spreadNodes = "";
			string spreadTemplate = _readFile(this.ContentSpreadTemplate);
			// Loop all photos
			var photos = System.IO.Directory.EnumerateFiles(this.PhotoDir).ToList();
			for(int i = 0; i < photos.Count(); i += 2) {
				// Init
				string spreadId = this.ContentSpreadName+i;
				string photo1 = photos[i];
				string photo2 = null;
				string page1Id = this.ContentSpreadName+i;
				string page2Id = this.ContentSpreadName+(i+1);
				if(photos.Count() > i+1) photo2 = photos[i+1];
				// Rewrite photo file path
				photo1 = "file:"+photo1;
				photo2 = "file:"+photo2;
				// Create spread file
				string spreadFile = spreadTemplate;
				spreadFile = spreadFile.Replace("[PHOTO1]",photo1);
				spreadFile = spreadFile.Replace("[PHOTO2]",photo2);
				spreadFile = spreadFile.Replace("[PAGE1_ID]",page1Id);
				spreadFile = spreadFile.Replace("[PAGE2_ID]",page2Id);
				spreadFile = spreadFile.Replace("[SPREAD_ID]",spreadId);
				// Write
				_writeFile(this.ContentSpreadTemplate.Replace(".xml","")+i+".xml",spreadFile);
				// Register
				var spreadFileName = "Spread_" + this.ContentSpreadName +i+ ".xml";
				spreadNodes += "\n" + spreadNodeTemplate.Replace("[FILE]",spreadFileName);
			}
			// Add varialbes
			_variables["[SPREADS]"] = spreadNodes;
		}

		public void ReplaceAllVariables() {
			_debug();
			// Loop all files
			_replaceVariablesInDir(this.WorkingFileDir);
		}

		public void ZipTemplate() {
			_debug();
			string outFile = this.OutputFile;
			string inDir = this.WorkingFileDir;
			_forceDelete(outFile);
			System.IO.Compression.ZipFile.CreateFromDirectory(inDir, outFile);
		}

		public void Cleanup() {

		}

		private void _copyDir(string sourcePath, string targetPath) {
			_debug(sourcePath, targetPath);
			if (!System.IO.Directory.Exists(targetPath)) {
				System.IO.Directory.CreateDirectory(targetPath);
			}
			foreach(var subdir in System.IO.Directory.EnumerateDirectories(sourcePath)) {
				string dirName = subdir.Replace(sourcePath, "");
				string subTargetPath = targetPath + dirName;
				_copyDir(subdir,subTargetPath);
			}
			string[] files = System.IO.Directory.GetFiles(sourcePath);
			foreach (string s in files) {
				var fileName = System.IO.Path.GetFileName(s);
				var destFile = System.IO.Path.Combine(targetPath, fileName);
				System.IO.File.Copy(s, destFile, true);
			}
		}

		private void _debug(string cat = null, string val = null) {
			if(!this.Verbose) return;
			string method = new System.Diagnostics.StackFrame(1, true).GetMethod().Name;
			if(val != null) {
				Console.WriteLine(method + ": "+cat + ": " + val);
			} else if(cat != null) {
				Console.WriteLine(method + ": "+cat);
			} else {
				Console.WriteLine(method);
			}
		}

		private void _replaceVariablesInDir(string dir) {
			_debug(dir);
			// Recurse dirs
			foreach(var subdir in System.IO.Directory.EnumerateDirectories(dir)) {
				_replaceVariablesInDir(subdir);
			}
			// Process files
			foreach(var file in System.IO.Directory.EnumerateFiles(dir)) {
				_replaceVariablesInFile(file);
			}
		}

		private string _readFile(string file) {
			var sr = new System.IO.StreamReader(file);
			string contents = sr.ReadToEnd();
			sr.Close();
			return contents;
		}

		private void _writeFile(string file, string contents) {
			var sw = new System.IO.StreamWriter(file);
			sw.Write(contents);
			sw.Close();
		}

		private void _replaceVariablesInFile(string file) {
			if(!file.EndsWith(".xml")) return;
			_debug(file);
			// Load file
			string contents = _readFile(file);
			// Replace
			foreach(var key in _variables.Keys) {
				contents = contents.Replace(key, _variables[key]);
			}
			// Save
			_writeFile(file, contents);
		}

		private void _processExec(string process, string args) {
			
		}

		private void _forceDelete(string file) {
			try {
				System.IO.File.Delete(file);
			} catch(Exception e) {}
			try {
				System.IO.Directory.Delete(file,true);
			} catch(Exception e) {}
		}
	}
}
