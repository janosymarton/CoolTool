using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

// This file contains the different request and response object definition that are used to 
// serialize to and from JSON expressions. 

namespace CoolTool
{
    class LoginRequest
    {
        public string username { get; set; }
        public string password { get; set; }
        public int LoginMode = 0;

    }

    public class LoginResponse
    {
        public string Name { get; set; }
        public string Sid { get; set; }
        public string AccessToken { get; set; }
    }
}

namespace ResourcesAPI_TMLookup
{
    public class SegmentObj
    {
        public string ContextID { get; set; }
        public string FollowingSegment { get; set; }
        public string PrecedingSegment { get; set; }
        public string Segment { get; set; }
    }

    public class Options
    {
        public bool AdjustFuzzyMatches { get; set; }
        public int InlineTagStrictness { get; set; }
        public int MatchThreshold { get; set; }
        public bool OnlyBest { get; set; }
        public bool OnlyUnambiguous { get; set; }
        public bool ReverseLookup { get; set; }
    }

    public class TMLookupRequest
    {
        public List<SegmentObj> Segments { get; set; }
        public Options Options { get; set; }
        public TMLookupRequest()
        {
            Segments = new List<SegmentObj>();
        }
    }

    public class CustomMeta
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public class TransUnit
    {
        public string Client { get; set; }
        public string ContextID { get; set; }
        public string Created { get; set; }
        public string Creator { get; set; }
        public string Document { get; set; }
        public string Domain { get; set; }
        public string FollowingSegment { get; set; }
        public int Key { get; set; }
        public string Modified { get; set; }
        public string Modifier { get; set; }
        public string PrecedingSegment { get; set; }
        public string Project { get; set; }
        public string SourceSegment { get; set; }
        public string Subject { get; set; }
        public string TargetSegment { get; set; }
        public List<CustomMeta> CustomMetas { get; set; }
        public TransUnit()
        {
            CustomMetas = new List<CustomMeta>();
        }
    }

    public class TMHit
    {
        public int MatchRate { get; set; }
        public TransUnit TransUnit { get; set; }
        public TMHit()
        {
            TransUnit = new TransUnit();
        }
    }

    public class Result
    {
        public List<TMHit> TMHits { get; set; }
        public Result()
        {
            TMHits = new List<TMHit>();
        }
    }

    public class TMLookupResponse
    {
        public List<Result> Result { get; set; }
        public TMLookupResponse()
        {
            Result = new List<Result>();
        }
    }

}
namespace ResourcesAPI_TBLookup
{
    public class TBLookupRequest
    {
        public string SourceLanguage { get; set; }
        public string TargetLanguage { get; set; }
        public List<string> Segments { get; set; }
        public TBLookupRequest()
        {
            Segments = new List<string>();
        }
    }

    public class Result
    {
        public List<List<TBLookupResult>> TBHits { get; set; }
        public Result()
        {
            TBHits = new List<List<TBLookupResult>>();
        }
    }

    public class TBLookupResponse
    {
        public List<Result> Result { get; set; }
        public TBLookupResponse()
        {
            Result = new List<Result>();
        }
    }

    public class CustomMeta
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public class TermItem
    {
        public int CaseSense { get; set; }
        public string Example { get; set; }
        public string GrammarGender { get; set; }
        public string GrammarNumber { get; set; }
        public string GrammarPartOfSpeech { get; set; }
        public int Id { get; set; }
        public bool IsForbidden { get; set; }
        public int PartialMatch { get; set; }
        public List<int> PrefixBoundaries { get; set; }
        public string Text { get; set; }
        public string WildText { get; set; }
        public List<CustomMeta> CustomMetas { get; set; }
        public TermItem()
        {
            CustomMetas = new List<CustomMeta>();
            PrefixBoundaries = new List<int>();
        }

    }


    public class LanguageObj
    {
        public string Language { get; set; }
        public string Definition { get; set; }
        public int Id { get; set; }
        public bool NeedsModeration { get; set; }
        public List<TermItem> TermItems { get; set; }
        public List<CustomMeta> CustomMetas { get; set; }
        public LanguageObj()
        {
            TermItems = new List<TermItem>();
            CustomMetas = new List<CustomMeta>();
        }

    }

    public class Entry
    {
        public int Id { get; set; }
        public string Client { get; set; }
        public string Created { get; set; }
        public string Creator { get; set; }
        public string Domain { get; set; }
        public string ImageCaption { get; set; }
        public string Image { get; set; }
        public List<LanguageObj> Languages { get; set; }
        public string Modified { get; set; }
        public string Modifier { get; set; }
        public string Note { get; set; }
        public string Project { get; set; }
        public string Subject { get; set; }
        public List<CustomMeta> CustomMetas { get; set; }

        public Entry()
        {
            Languages = new List<LanguageObj>();
            CustomMetas = new List<CustomMeta>();
        }
    }

    public class TBLookupResult
    {
        public int LengthInSegment { get; set; }
        public int MatchRate { get; set; }
        public string SourceLang { get; set; }
        public string SourceTerm { get; set; }
        public int SourceTermIndex { get; set; }
        public int StartPosInSegment { get; set; }
        public string TargetLang { get; set; }
        public string TargetTerm { get; set; }
        public int TargetTermIndex { get; set; }
        public Entry Entry { get; set; }
        public TBLookupResult()
        {
            Entry = new Entry();
        }
    }
}

namespace ResourcesAPI_TMList
{
    public class TMListResponse
    {
        public int AccessLevel { get; set; }
        public string Client { get; set; }
        public string Domain { get; set; }
        public string Subject { get; set; }
        public string Project { get; set; }
        public int NumEntries { get; set; }
        public string FriendlyName { get; set; }
        public string SourceLangCode { get; set; }
        public string TargetLangCode { get; set; }
        public string TMGuid { get; set; }
        public string TMOwner { get; set; }
    }

}

namespace ResourcesAPI_TBList
{
    public class TBListResponse
    {

        public int AccessLevel { get; set; }
        public string Client { get; set; }
        public string Domain { get; set; }
        public string FriendlyName { get; set; }
        public List<string> Languages { get; set; }
        public int NumEntries { get; set; }
        public string Project { get; set; }
        public string Subject { get; set; }
        public string TBGuid { get; set; }
        public string TBOwner { get; set; }
        public TBListResponse()
        {
            Languages = new List<string>();
        }
    }
}

namespace ResourcesAPI_TBEntry
{
    public class TBEntry
    {
        public int Id { get; set; }
        public string Client { get; set; }
        public string Created { get; set; }
        public string Creator { get; set; }
        public string Domain { get; set; }
        public string ImageCaption { get; set; }
        public string Image { get; set; }
        public List<Language> Languages { get; set; }
        public string Modified { get; set; }
        public string Modifier { get; set; }
        public string Note { get; set; }
        public string Project { get; set; }
        public string Subject { get; set; }
        public List<CustomMeta> CustomMetas { get; set; }

    }


    public class Language
    {
        public string language { get; set; }
        public string Definition { get; set; }
        public int Id { get; set; }
        public bool NeedsModeration { get; set; }
        public List<TermItem> TermItems { get; set; }
        public List<CustomMeta> CustomMetas { get; set; }
    }


    public class TermItem
    {
        public int CaseSense { get; set; }
        public string Example { get; set; }
        public string GrammarGender { get; set; }
        public string GrammarNumber { get; set; }
        public string GrammarPartOfSpeech { get; set; }
        public int Id { get; set; }
        public bool IsForbidden { get; set; }
        public int PartialMatch { get; set; }
        public List<int> PrefixBoundaries { get; set; }
        public string Text { get; set; }
        public string WildText { get; set; }
        public List<CustomMeta> CustomMetas { get; set; }
    }


    public class CustomMeta
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }


}

