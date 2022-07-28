using System;
using AutoFixture;
using AutoFixture.AutoMoq;
using AutoFixture.Kernel;
using Moq;

namespace Flexcel.Tests.Infrastructure;

public static class FixtureExtensions
{
    public static void ConfigureMoqIntegration(this IFixture fixture)
    {
        var specimenBuilder = new Postprocessor(new MethodInvoker(new ModestConstructorQuery()),
                                                new MockFiller());
        fixture.Customizations.Add(new FilteringSpecimenBuilder(specimenBuilder, new MockSpecification()));
        fixture.Customize(new AutoMoqCustomization { ConfigureMembers = true });
    }

    public static Mock<TParameter> FreezeMock<TParameter>(this IFixture fixture) where TParameter : class
    {
        if (fixture == null) throw new ArgumentNullException(nameof(fixture));
        
        var mock = fixture.Create<Mock<TParameter>>();
        fixture.Inject(mock.Object);
        return mock;
    }
}