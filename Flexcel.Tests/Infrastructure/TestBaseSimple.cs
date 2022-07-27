using System.IO;
using System.Threading.Tasks;
using AutoFixture;
using NUnit.Framework;

namespace Flexcel.Tests.Infrastructure;

public abstract class TestBaseSimple
{
    protected IFixture fixture;
        
    [SetUp]
    public virtual void SetUp()
    {
        fixture = FixtureCreator.Create();
        SetupAsync().GetAwaiter().GetResult();
    }

    public virtual Task SetupAsync()
    {
        return Task.CompletedTask;
    }

    [TearDown]
    public virtual void TearDown()
    {
    }

    [OneTimeSetUp]
    public virtual void FixtureSetUp()
    {
        Directory.SetCurrentDirectory(TestContext.CurrentContext.TestDirectory);
        FixtureSetupAsync().GetAwaiter().GetResult();
    }
        
    public virtual Task FixtureSetupAsync() => Task.CompletedTask;

    [OneTimeTearDown]
    public virtual void FixtureTearDown()
    {
    }
}