using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTSafeCommands.Utilities
{
    /// <summary>
    /// Represents a non-generic operation result.
    /// </summary>
    public class BooleanResult
    {
        public static readonly BooleanResult DefaultResult =
            new BooleanResult();

        /// <summary>
        /// Creates a successful result.
        /// </summary>
        public static BooleanResult SuccessResult(
            string message)
        {
            return new BooleanResult(true, message);
        }

        /// <summary>
        /// Creates a failed result.
        /// </summary>
        public static BooleanResult FailResult(
            string error)
        {
            return new BooleanResult(false, error);
        }

        private BooleanResult(
            bool success = true,
            string message = null)
        {
            Success = success;
            Message = message;
        }

        /// <summary>
        /// Copies failure details to a generic result.
        /// </summary>
        public BooleanResult<TReturn>
            CopyFailureDetailsTo<TReturn>()
        {
            return BooleanResult<TReturn>.FailResult(
                Message);
        }

        public bool Success { get; protected set; }

        public string Message { get; protected set; }

        public bool IsSuccess => Success;
    }

    /// <summary>
    /// Represents a generic operation result.
    /// </summary>
    public class BooleanResult<T>
    {
        /// <summary>
        /// Creates a successful result.
        /// </summary>
        public static BooleanResult<T> SuccessResult(
            T result)
        {
            return new BooleanResult<T>
            {
                Success = true,
                Result = result
            };
        }

        /// <summary>
        /// Creates a failed result.
        /// </summary>
        public static BooleanResult<T> FailResult(
            string error)
        {
            return new BooleanResult<T>
            {
                Success = false,
                Message = error
            };
        }

        /// <summary>
        /// Copies failure details to another generic result.
        /// </summary>
        public BooleanResult<TReturn>
            CopyFailureDetailsTo<TReturn>()
        {
            return new BooleanResult<TReturn>
            {
                Success = Success,
                Message = Message
            };
        }

        /// <summary>
        /// Copies failure details to a non-generic result.
        /// </summary>
        public BooleanResult CopyFailureDetailsTo()
        {
            return BooleanResult.FailResult(Message);
        }

        public bool Success { get; protected set; }

        public string Message { get; protected set; }

        public T Result { get; protected set; }
    }
}