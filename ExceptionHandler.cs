using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommonUtil;

namespace NGPReportManager
{
    /// <summary>
    /// 本类主要用来处理异常
    /// 当捕获到异常时重复调用，直到调用成功为止
    /// </summary>
    class ExceptionHandler
    {
        //下面这6组方法是针对异常：被呼叫方拒绝接收呼叫。 (异常来自 HRESULT:0x80010001 (RPC_E_CALL_REJECTED))
        //异常类型：System.Runtime.InteropServices.COMException
        //ErrorCode = -2147418111
        public static void RunWithOutRejected(Action action)
        {
            bool hasException;
            do
            {
                try
                {
                    action();
                    hasException = false;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    if (e.ErrorCode == -2147418111)
                    {
                        TraceLogManager.Instance.WriteLogFile(string.Format("RunWithOutRejected被调用,处理{0}“被呼叫方拒绝接受呼叫”异常",action.Method.ToString()));
                        hasException = true;
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            } while (hasException);
        }

        public static void RunWithOutRejected<T>(Action<T> action, T t)
        {
            bool hasException;
            do
            {
                try
                {
                    action(t);
                    hasException = false;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    if (e.ErrorCode == -2147418111)
                    {
                        TraceLogManager.Instance.WriteLogFile(string.Format("RunWithOutRejected被调用,处理 {0} “被呼叫方拒绝接受呼叫”异常", action.Method.ToString()));
                        hasException = true;
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            } while (hasException);
        }

        public static void RunWithOutRejected<T1, T2, T3>(Action<T1, T2, T3> action, T1 t1, T2 t2, T3 t3)
        {
            bool hasException;
            do
            {
                try
                {
                    action(t1, t2, t3);
                    hasException = false;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    if (e.ErrorCode == -2147418111)
                    {
                        TraceLogManager.Instance.WriteLogFile(string.Format("RunWithOutRejected被调用,处理{0}“被呼叫方拒绝接受呼叫”异常", action.Method.ToString()));
                        hasException = true;
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            } while (hasException);
        }

        public static T RunWithOutRejected<T>(Func<T> func)
        {
            var result = default(T);
            bool hasException;
            do
            {
                try
                {
                    result = func();
                    hasException = false;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    if (e.ErrorCode == -2147418111)
                    {
                        TraceLogManager.Instance.WriteLogFile(string.Format("RunWithOutRejected<T>(Func<T> func)被调用,处理 {0}“被呼叫方拒绝接受呼叫”异常",func.ToString()));                        
                        hasException = true;
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            } while (hasException);
            return result;
        }

        public static TResult RunWithOutRejected<T, TResult>(Func<T, TResult> func,T t)
        {
            var result = default(TResult);
            bool hasException;
            do
            {
                try

                {
                    result = func(t);
                    hasException = false;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    if (e.ErrorCode == -2147418111)
                    {
                        TraceLogManager.Instance.WriteLogFile(string.Format("RunWithOutRejected被调用,处理{0}“被呼叫方拒绝接受呼叫”异常",func.Method.ToString()));
                        hasException = true;
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            } while (hasException);
            return result;
        }

        public static TResult RunWithOutRejected<T1, T2, TResult>(Func<T1, T2, TResult> func, T1 t1, T2 t2)
        {
            var result = default(TResult);
            bool hasException;
            do
            {
                try
                {
                    result = func(t1, t2);
                    hasException = false;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    if (e.ErrorCode == -2147418111)
                    {
                        TraceLogManager.Instance.WriteLogFile(string.Format("RunWithOutRejected被调用,处理{0}“被呼叫方拒绝接受呼叫”异常", func.Method.ToString()));
                        hasException = true;
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            } while (hasException);
            return result;
        }

        //TODO:用来进行处理剪切板为空的异常。目前在复制前先清空了剪切板，这个类型的异常没有出现，暂时未实现。 
        //异常类型：System.Runtime.InteropServices.COMException
        //ErrorCode = -2146823683
         public static void RunWithClipboardEmpty(Action action){}
       
       
    }
}
