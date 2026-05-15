using System;
using System.Collections.Generic;
using System.Text;

namespace Sample_exercise.Utilities
{
    public class BooleanResult
    {
        public static readonly BooleanResult DefaultResult =
            new BooleanResult();

        public static BooleanResult SuccessResult(
            string message)
        {
            return new BooleanResult(
                true,
                message);
        }

        public static BooleanResult FailResult(
            string error)
        {
            return new BooleanResult(
                false,
                error);
        }

        private BooleanResult(
            bool success = true,
            string message = null)
        {
            Success = success;
            Message = message;
        }

        public bool Success { get; protected set; }

        public string Message { get; protected set; }
    }

    public class BooleanResult<T>
    {
        public static BooleanResult<T> SuccessResult(
            T result)
        {
            return new BooleanResult<T>
            {
                Success = true,
                Result = result
            };
        }

        public static BooleanResult<T> FailResult(
            string message)
        {
            return new BooleanResult<T>
            {
                Success = false,
                Message = message
            };
        }

        public bool Success { get; protected set; }

        public string Message { get; protected set; }

        public T Result { get; protected set; }
    }
}
