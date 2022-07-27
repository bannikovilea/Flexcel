using AutoFixture.Kernel;
using Moq;

namespace Flexcel.Tests.Infrastructure;

internal class MockFiller : ISpecimenCommand
{
    public void Execute(object specimen, ISpecimenContext context)
    {
        var mock = specimen as Mock;
        mock.DefaultValue = DefaultValue.Empty;
    }
}