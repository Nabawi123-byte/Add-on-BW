using System.Collections.Generic;
using System.Linq;
using System.Text;
using MasterBW.Core.Tools;

namespace MasterBW.Core.Ai
{
    /// <summary>
    /// Membangun grammar GBNF (llama.cpp) yang MEMAKSA model kecil mengeluarkan JSON rencana aksi yang valid
    /// dan hanya memakai nama aksi yang ada di katalog. Ini kunci keandalan model 1,5–4B.
    /// </summary>
    public static class PlanGrammar
    {
        public static string Build(IEnumerable<ToolDef> tools, int maxSteps = 12)
        {
            var names = tools.Select(t => t.Name).Distinct().ToList();
            var sb = new StringBuilder();
            sb.AppendLine("root ::= \"{\" ws \"\\\"message\\\"\" ws \":\" ws string \",\" ws \"\\\"steps\\\"\" ws \":\" ws steps ws \"}\"");
            sb.Append("steps ::= \"[\" ws \"]\" | \"[\" ws step");
            sb.Append(" (\",\" ws step){0,").Append(maxSteps - 1).AppendLine("} ws \"]\"");
            sb.AppendLine("step ::= \"{\" ws \"\\\"action\\\"\" ws \":\" ws action \",\" ws \"\\\"args\\\"\" ws \":\" ws object ws \"}\" ws");
            sb.Append("action ::= ");
            sb.AppendLine(string.Join(" | ", names.Select(n => "\"\\\"" + n + "\\\"\"")));
            sb.AppendLine(CommonJson);
            return sb.ToString();
        }

        /// <summary>Grammar JSON umum (objek apa saja).</summary>
        public static string AnyObject => "root ::= object\n" + CommonJson;

        private const string CommonJson =
@"value ::= object | array | string | number | (""true"" | ""false"" | ""null"") ws
object ::= ""{"" ws ( string "":"" ws value ( "","" ws string "":"" ws value )* )? ""}"" ws
array ::= ""["" ws ( value ( "","" ws value )* )? ""]"" ws
string ::= ""\"""" char* ""\"""" ws
char ::= [^""\\\x7F\x00-\x1F] | ""\\"" ([""\\/bfnrt] | ""u"" [0-9a-fA-F] [0-9a-fA-F] [0-9a-fA-F] [0-9a-fA-F])
number ::= ""-""? ([0-9] | [1-9] [0-9]{0,15}) (""."" [0-9]{1,10})? ([eE] [-+]? [0-9]{1,3})? ws
ws ::= | "" "" | ""\n"" [ \t]{0,20}";
    }
}
