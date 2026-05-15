using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace QueenMasterVisio.DeviceControl.Rules
{
    internal class PlanTraccers : IRule
    {
            public string Name => "Размер фигуры";
            public ResultSeverity Severity { get; set; } = ResultSeverity.Error;

            public IEnumerable<ValidationResult> Validate(Visio.Shape shape, ValidationContext context)
            {
                var expectedWidth = GetExpectedWidth(shape); // логика из бизнес-правил
                var actualWidth = shape.CellsU["Width"].ResultIU; // в дюймах или мм

                if (Math.Abs(actualWidth - expectedWidth) > context.ToleranceMm / 25.4) // IU - дюймы
                {
                    yield return new ValidationResult
                    {
                        RuleName = Name,
                        TargetShape = shape,
                        Severity = this.Severity,
                        Message = $"Ожидаемая ширина: {expectedWidth * 25.4:F1} мм, фактическая: {actualWidth * 25.4:F1} мм",
                        ExpectedValue = expectedWidth,
                        ActualValue = actualWidth
                    };
                }
            }
        
    }
}
