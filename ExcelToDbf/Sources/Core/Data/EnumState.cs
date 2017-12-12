using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToDbf.Sources.Core.Data
{
    public enum EnumState : byte
    {
        CHOOSE_FILES = 1,
        VIEW_LOG = 2
    }

    public static class EnumStateExtension
    {
        public static EnumState Next(this EnumState state)
        {
            switch (state)
            {
                case EnumState.CHOOSE_FILES:
                    return EnumState.VIEW_LOG;
                case EnumState.VIEW_LOG:
                    return EnumState.CHOOSE_FILES;
            }
            // Unreachable code?
            throw new NotImplementedException($"Для значения '{state}' enum '{nameof(EnumState)}' не предусмотрено перехода в следующее состояние!");
        }
    }
}
