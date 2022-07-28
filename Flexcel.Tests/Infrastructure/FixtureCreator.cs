using AutoFixture;

namespace Flexcel.Tests.Infrastructure;

internal static class FixtureCreator
{
    public static Fixture Create()
    {
        var fixture = new Fixture();
        fixture.ConfigureMoqIntegration();
        fixture.Customize(new SupportMutableValueTypesCustomization());
            
        return fixture;
    }
}