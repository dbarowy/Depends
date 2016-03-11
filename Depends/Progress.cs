using System;
using System.Threading;

namespace Depends
{
    public delegate void ProgressBarIncrementer();

    public class Progress
    {
        private long _total = 0;
        private ProgressBarIncrementer _progBarIncr;

        public Progress(ProgressBarIncrementer progBarIncrement)
        {
            _progBarIncr = progBarIncrement;
        }

        public long Total
        {
            get { return _total; }
            set { _total = value; }
        }

        public long UpdateEvery
        {
            get { return Math.Max(1L, _total / 100L); }
        }

        public void IncrementCounter()
        {
            _progBarIncr();
        }
    }
}
