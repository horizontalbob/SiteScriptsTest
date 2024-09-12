using System.Collections.Generic;
using Newtonsoft.Json;

namespace SiteScriptTestFramework
{
    public class SiteScript
    {
        [JsonProperty("$schema")]
        public string schema { get; set; }
        public List<Action> actions { get; set; } = new List<Action>();
        public BindData bindData {  get; set; }
        public int version => 1;
    }

    public class Action
    {
        public string verb { get; set; }
        public string themeName { get; set; }
        public string listName { get; set; }
        public int templateType { get; set; }
        public List<Subaction> subactions { get; set; } = new List<Subaction>();
    }

    public class Subaction
    {
        public string verb { get; set; }
        public string description { get; set; }
        public string internalName { get; set; }
        public string fieldType { get; set; }
        public string displayName { get; set; }
        public bool isRequired { get; set; }
        public bool addToDefaultView { get; set; }
    }

    public class BindData
    {

    }

}
