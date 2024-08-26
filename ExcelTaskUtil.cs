using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace TestTaskAsync
{
    // Helper class to wrap a Task in an Observable - allowing one Subscriber.
    internal class ExcelTaskObservable<TResult> : IExcelObservable
    {
        readonly Task<TResult> _task;
        readonly CancellationTokenSource _cts;

        public ExcelTaskObservable(Task<TResult> task)
        {
            _task = task;
        }

        public ExcelTaskObservable(Task<TResult> task, CancellationTokenSource cts)
            : this(task)
        {
            _cts = cts;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            // Start with a disposable that does nothing
            // Possibly set to a CancellationDisposable later
            IDisposable disp = DefaultDisposable.Instance;

            switch (_task.Status)
            {
                case TaskStatus.RanToCompletion:
                    observer.OnNext(_task.Result);
                    observer.OnCompleted();
                    break;
                case TaskStatus.Faulted:
                    observer.OnError(_task.Exception.InnerException);
                    break;
                case TaskStatus.Canceled:
                    observer.OnError(new TaskCanceledException(_task));
                    break;
                default:
                    var task = _task;
                    // OK - the Task has not completed synchronously
                    // First set up a continuation that will suppress Cancel after the Task completes
                    if (_cts != null)
                    {
                        var cancelDisp = new CancellationDisposable(_cts);
                        task = _task.ContinueWith(t =>
                        {
                            cancelDisp.SuppressCancel();
                            return t;
                        }).Unwrap();

                        // Then this will be the IDisposable we return from Subscribe
                        disp = cancelDisp;
                    }
                    // And handle the Task completion
                    task.ContinueWith(t =>
                    {
                        switch (t.Status)
                        {
                            case TaskStatus.RanToCompletion:
                                observer.OnNext(t.Result);
                                observer.OnCompleted();
                                break;
                            case TaskStatus.Faulted:
                                observer.OnError(t.Exception.InnerException);
                                break;
                            case TaskStatus.Canceled:
                                observer.OnError(new TaskCanceledException(t));
                                break;
                        }
                    });
                    break;
            }

            return disp;
        }

        sealed class DefaultDisposable : IDisposable
        {
            public static readonly DefaultDisposable Instance = new DefaultDisposable();

            // Prevent external instantiation
            DefaultDisposable()
            {
            }

            public void Dispose()
            {
                // no op
            }
        }

        sealed class CancellationDisposable : IDisposable
        {
            bool _suppress;
            readonly CancellationTokenSource _cts;
            public CancellationDisposable(CancellationTokenSource cts)
            {
                if (cts == null)
                {
                    throw new ArgumentNullException("cts");
                }

                _cts = cts;
            }

            public CancellationDisposable()
                : this(new CancellationTokenSource())
            {
            }

            public void SuppressCancel()
            {
                _suppress = true;
            }

            public CancellationToken Token
            {
                get { return _cts.Token; }
            }

            public void Dispose()
            {
                if (!_suppress) _cts.Cancel();
                _cts.Dispose();  // Not really needed...
            }
        }
    }

    // An IExcelObservable that wraps an IObservable
    internal class ExcelObservable<T> : IExcelObservable
    {
        readonly IObservable<T> _observable;

        public ExcelObservable(IObservable<T> observable)
        {
            _observable = observable;
        }

        public IDisposable Subscribe(IExcelObserver excelObserver)
        {
            var observer = new AnonymousObserver<T>(value => excelObserver.OnNext(value), excelObserver.OnError, excelObserver.OnCompleted);
            return _observable.Subscribe(observer);
        }

        // An IObserver that forwards the inputs to given methods.
        class AnonymousObserver<OT> : IObserver<OT>
        {
            readonly Action<OT> _onNext;
            readonly Action<Exception> _onError;
            readonly Action _onCompleted;

            public AnonymousObserver(Action<OT> onNext, Action<Exception> onError, Action onCompleted)
            {
                if (onNext == null)
                {
                    throw new ArgumentNullException("onNext");
                }
                if (onError == null)
                {
                    throw new ArgumentNullException("onError");
                }
                if (onCompleted == null)
                {
                    throw new ArgumentNullException("onCompleted");
                }
                _onNext = onNext;
                _onError = onError;
                _onCompleted = onCompleted;
            }

            public void OnNext(OT value)
            {
                _onNext(value);
            }

            public void OnError(Exception error)
            {
                _onError(error);
            }

            public void OnCompleted()
            {
                _onCompleted();
            }
        }
    }

    internal class ExcelTaskUtil
    {
        public static object RunTask<TResult>(string callerFunctionName, object callerParameters, Func<Task<TResult>> taskSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, delegate
            {
                var task = taskSource();
                return new ExcelTaskObservable<TResult>(task);
            });
        }
    }
}
