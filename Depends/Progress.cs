using System;
using System.Threading;

namespace Depends
{
    public delegate void ProgressBarIncrementer();
    public delegate void ProgressBarReset();

    public class Progress
    {
        private volatile bool _cancelled = false;
        private long _total = 0;
        private ProgressBarIncrementer _progBarIncr;
        private ProgressBarReset _progBarReset;
        private long _workMultiplier = 1;

        public static Progress NOPProgress()
        {
            ProgressBarIncrementer pbi = () => { return; };
            ProgressBarReset pbr = () => { return; };
            return new Progress(pbi, pbr, 1L);
        }

        public Progress(ProgressBarIncrementer progBarIncrement, ProgressBarReset progBarReset, long workMultiplier)
        {
            _progBarIncr = progBarIncrement;
            _progBarReset = progBarReset;
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

        public void Cancel()
        {
            _cancelled = true;
        }

        public bool IsCancelled()
        {
            return _cancelled;
        }

        public void Reset()
        {
            _progBarReset();
        }
    }
}
