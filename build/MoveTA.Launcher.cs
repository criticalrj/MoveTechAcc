// Move TA - launcher portatil (arquivo unico).
// Leva Move-TA.ps1 e usp_IPS_MoveTA.sql como recursos embutidos, extrai para %TEMP%\MoveTA_<versao>
// e abre o app com o PowerShell 5.1 do Windows (sem console). Logs ficam na pasta "logs" ao lado do .exe.
// Uso:  MoveTA.exe                -> abre o app
//       MoveTA.exe /debug         -> abre com console visivel (para ver erros)
//       MoveTA.exe -Lang en ...   -> demais parametros vao direto para o Move-TA.ps1
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

[assembly: AssemblyTitle("Move TA")]
[assembly: AssemblyProduct("Move TA - STAR IPS 3.62 Technical Accounts")]
[assembly: AssemblyDescription("Realocar, criar e renomear Technical Accounts no STAR IPS 3.62")]
[assembly: AssemblyVersion("2.2.0.0")]
[assembly: AssemblyFileVersion("2.2.0.0")]
[assembly: AssemblyInformationalVersion("2.2")]
[assembly: AssemblyCompany("Flávio Miguel")]
[assembly: AssemblyCopyright("© 2026 Flávio Miguel")]

static class Program
{
    const string Version = "2.2";

    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            string work = Path.Combine(Path.GetTempPath(), "MoveTA_" + Version);
            Directory.CreateDirectory(work);
            string script = Extract("Move-TA.ps1", work);
            Extract("usp_IPS_MoveTA.sql", work);

            string logs = Path.Combine(exeDir, "logs");
            try { Directory.CreateDirectory(logs); }
            catch { logs = Path.Combine(work, "logs"); }   // pendrive somente leitura: usa o TEMP

            bool debug = false;
            var extra = new StringBuilder();
            foreach (string a in args)
            {
                if (a.Equals("/debug", StringComparison.OrdinalIgnoreCase) || a.Equals("-debug", StringComparison.OrdinalIgnoreCase)) { debug = true; continue; }
                extra.Append(' ').Append(Quote(a));
            }

            string ps = Path.Combine(Environment.SystemDirectory, @"WindowsPowerShell\v1.0\powershell.exe");
            if (!File.Exists(ps)) { Fail("PowerShell 5.1 não encontrado em " + ps); return 2; }

            string psArgs = "-NoProfile -ExecutionPolicy Bypass -STA " + (debug ? "-NoExit " : "-WindowStyle Hidden ")
                          + "-File " + Quote(script) + " -LogDir " + Quote(logs) + extra.ToString();
            var psi = new ProcessStartInfo(ps, psArgs);
            psi.UseShellExecute = false;
            psi.CreateNoWindow = !debug;
            psi.WorkingDirectory = work;
            Process p = Process.Start(psi);

            // se o app morrer logo no início, avisa (em vez de "não abriu nada")
            if (!debug && p.WaitForExit(6000) && p.ExitCode != 0)
            {
                Fail("O Move TA fechou ao iniciar (código " + p.ExitCode + ").\n\n" +
                     "Rode \"MoveTA.exe /debug\" para ver o erro.\n" +
                     "Causas comuns: política do Windows bloqueando scripts PowerShell (GPO AllSigned) ou .NET/WPF ausente.");
                return p.ExitCode;
            }
            return 0;
        }
        catch (Exception ex) { Fail(ex.Message); return 1; }
    }

    static string Extract(string name, string dir)
    {
        string target = Path.Combine(dir, name);
        using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
        {
            if (s == null) throw new Exception("Recurso embutido não encontrado: " + name);
            byte[] data = new byte[s.Length];
            int read = 0; while (read < data.Length) { int n = s.Read(data, read, data.Length - read); if (n <= 0) break; read += n; }
            if (!File.Exists(target) || !SameBytes(File.ReadAllBytes(target), data)) File.WriteAllBytes(target, data);
        }
        return target;
    }

    static bool SameBytes(byte[] a, byte[] b)
    {
        if (a.Length != b.Length) return false;
        for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
        return true;
    }

    static string Quote(string s) { return "\"" + s.Replace("\"", "\\\"") + "\""; }

    static void Fail(string msg) { MessageBox.Show(msg, "Move TA", MessageBoxButtons.OK, MessageBoxIcon.Error); }
}
