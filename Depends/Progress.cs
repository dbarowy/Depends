using System;
using System.Threading;

namespace Depends
{
    public delegate void ProgressBarIncrementer();

    public class Progress
    {
        private long _total = 0;
        private ProgressBarIncrementer _progBarIncr;
        private long _workMultiplier = 1;

        public static Progress NOPProgress()
        {
            ProgressBarIncrementer pbi = () => { return; };
            return new Progress(pbi, 1L);
        }

        public Progress(ProgressBarIncrementer progBarIncrement, long workMultiplier)
        {
            _progBarIncr = progBarIncrement;
            _workMultiplier = workMultiplier;
        }

        public long TotalWorkUnits
        {
            get { return _total; }
            set { _total = value; }
        }

        public long UpdateEvery
        {
            get { return Math.Max(1L, _total / 100L / _workMultiplier); }
        }

        public void IncrementCounter()
        {
            _progBarIncr();
        }
    }
}
