using System.Collections;
using UnityEngine;
using UnityEngine.Networking;

namespace SNUPlugin
{
    class SNUWebRequest : MonoBehaviour
    {
        public static string staticGet(string uri) {
            var obj = new GameObject("SNUPlugin");
            var inst = obj.AddComponent<SNUWebRequest>();
            DestroyImmediate(obj);
            return inst.Get(uri);
        }

        public static string staticPost(string uri, string postdata)
        {
            var obj = new GameObject("SNUPlugin");
            var inst = obj.AddComponent<SNUWebRequest>();
            DestroyImmediate(obj);
            return inst.Post(uri, postdata);
        }

        private object getAsyncResult(IEnumerator func)
        {
            while (func.MoveNext()) ;
            return func.Current;
        }

        public string Post(string uri, string postdata)
        {
            return (string)getAsyncResult(this.__Post(uri, postdata));
        }

        private IEnumerator __Post(string uri, string postdata)
        {
            UnityWebRequest www = UnityWebRequest.Post(uri, postdata);
            yield return www.Send();
            while (!www.isDone)
                yield return null;
            if (www.isError)
                yield return www.error;
            else
                yield return www.downloadHandler.text;
        }

        public string Get(string uri)
        {
            return (string)getAsyncResult(__Get(uri));
        }

        private IEnumerator __Get(string uri)
        {
            UnityWebRequest www = UnityWebRequest.Get(uri);
            yield return www.Send(); 
            while (!www.isDone)
                yield return null;
            if (www.isError)
                yield return www.error;
            else
                yield return www.downloadHandler.text;
        }
    }
}
