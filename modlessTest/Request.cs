using System;
using System.Threading;

namespace modlessTest
{
    public enum RequestId : int
    {
        None = 0,
        Test = 1,
        TestTest = 2,
        gangpin = 3,
        readCAD = 4,
        CreateBeams = 5,
        deckslab = 6,
        deckSlab2 = 7
    }

    public class Request
    {
        private int m_request = (int)RequestId.None;

        public RequestId Take()
        {
            return (RequestId)Interlocked.Exchange(ref m_request, (int)RequestId.None);
        }

        public void Make(RequestId request)
        {
            Interlocked.Exchange(ref m_request, (int)request);
        }
    }
}
