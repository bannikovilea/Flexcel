using AutoFixture.Kernel;
using Moq;

namespace Flexcel.Tests.Infrastructure;

internal class MockSpecification : IRequestSpecification
{
    public bool IsSatisfiedBy(object request)
    {
        var requestType = request as Type;
        if (requestType == null)
            return false;

        return typeof(Mock).IsAssignableFrom(requestType);
    }
}