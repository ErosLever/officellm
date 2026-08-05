using OfficeMcpServer.Models;
using System.Collections.Generic;

namespace OfficeMcpServer.Tests;

public class InstanceRegistryTests
{
    [Fact]
    public void RegisterInstance_UsesProposedId_WhenValid()
    {
        var registry = new InstanceRegistry();
        var id1 = registry.RegisterInstance("PowerPoint", "deck1.pptx", "ppt_deck1_a3f2c1");
        var id2 = registry.RegisterInstance("PowerPoint", "deck2.pptx", "ppt_deck2_b7e409");

        Assert.Equal("ppt_deck1_a3f2c1", id1);
        Assert.Equal("ppt_deck2_b7e409", id2);
        Assert.NotEqual(id1, id2);
    }

    [Fact]
    public void RegisterInstance_AppearsInActiveInstances()
    {
        var registry = new InstanceRegistry();
        registry.RegisterInstance("PowerPoint", "test.pptx");

        var active = registry.GetActiveInstances();
        Assert.Single(active);
        Assert.Equal("PowerPoint", active[0].AppName);
        Assert.Equal("test.pptx", active[0].DocumentName);
    }

    [Fact]
    public void GetInstance_ReturnsInstance_WhenAlive()
    {
        var registry = new InstanceRegistry();
        var id = registry.RegisterInstance("PowerPoint", "deck.pptx");

        var instance = registry.GetInstance(id);
        Assert.NotNull(instance);
        Assert.Equal(id, instance.InstanceId);
    }

    [Fact]
    public void GetInstance_ReturnsNull_WhenNotFound()
    {
        var registry = new InstanceRegistry();
        var instance = registry.GetInstance("nonexistent");
        Assert.Null(instance);
    }

    [Fact]
    public void UpdateHeartbeat_UpdatesAppNameAndDocument()
    {
        var registry = new InstanceRegistry();
        var id = registry.RegisterInstance("PowerPoint", "old.pptx");

        registry.UpdateHeartbeat(id, "PowerPoint", "new.pptx");

        var instance = registry.GetInstance(id);
        Assert.NotNull(instance);
        Assert.Equal("new.pptx", instance.DocumentName);
    }

    [Fact]
    public void UpdateHeartbeat_DoesNothing_ForUnknownInstance()
    {
        var registry = new InstanceRegistry();
        // Should not throw
        registry.UpdateHeartbeat("nonexistent", "PowerPoint", "test.pptx");
    }

    [Fact]
    public void CleanupTimedOut_MarksStaleInstancesDead()
    {
        var registry = new InstanceRegistry();
        var id = registry.RegisterInstance("PowerPoint", "stale.pptx");

        // Manually expire the heartbeat by accessing the instance
        var instance = registry.GetInstance(id);
        Assert.NotNull(instance);
        // Simulate 60s passing by setting LastHeartbeat to the past
        instance.LastHeartbeat = DateTime.UtcNow.AddSeconds(-120);

        registry.CleanupTimedOut();

        // The instance should now be timed out and not returned
        var result = registry.GetInstance(id);
        Assert.Null(result);
    }

    [Fact]
    public void MultipleInstances_RegisterIndependently()
    {
        var registry = new InstanceRegistry();
        var id1 = registry.RegisterInstance("PowerPoint", "a.pptx");
        var id2 = registry.RegisterInstance("PowerPoint", "b.pptx");

        var active = registry.GetActiveInstances();
        Assert.Equal(2, active.Count);
        Assert.Contains(active, i => i.InstanceId == id1);
        Assert.Contains(active, i => i.InstanceId == id2);
    }

    [Fact]
    public void ActiveInstances_ExcludesTimedOut()
    {
        var registry = new InstanceRegistry();
        var id1 = registry.RegisterInstance("PowerPoint", "alive.pptx");
        var id2 = registry.RegisterInstance("PowerPoint", "dead.pptx");

        // Expire the second instance
        var dead = registry.GetInstance(id2);
        Assert.NotNull(dead);
        dead.LastHeartbeat = DateTime.UtcNow.AddSeconds(-120);

        var active = registry.GetActiveInstances();
        Assert.Single(active);
        Assert.Equal(id1, active[0].InstanceId);
    }

    [Fact]
    public void RegisterInstance_InvalidProposedId_FallsBackToRandomId()
    {
        var registry = new InstanceRegistry();
        // Invalid: contains uppercase, spaces, starts with digit
        var id1 = registry.RegisterInstance("PowerPoint", "deck.pptx", "Bad ID!");
        var id2 = registry.RegisterInstance("Word", "doc.docx", "1invalid");
        var id3 = registry.RegisterInstance("Excel", "sheet.xlsx", null);

        // All should fall back to something valid (office_ prefix + hex)
        Assert.All(new[] { id1, id2, id3 }, id =>
        {
            Assert.NotEmpty(id);
            Assert.Matches(@"^[a-z][a-z0-9_]+$", id);
        });
        // All distinct
        Assert.Equal(3, new HashSet<string> { id1, id2, id3 }.Count);
    }

    [Fact]
    public void RegisterInstance_SameProposedId_RefreshesHeartbeat()
    {
        var registry = new InstanceRegistry();
        var id1 = registry.RegisterInstance("Word", "doc.docx", "word_doc_abc123");
        var id2 = registry.RegisterInstance("Word", "doc.docx", "word_doc_abc123");

        Assert.Equal(id1, id2);
        Assert.Single(registry.GetActiveInstances());
    }

    [Fact]
    public void RegisterInstance_DifferentHosts_AllDistinct()
    {
        var registry = new InstanceRegistry();
        var ppt  = registry.RegisterInstance("PowerPoint", "deck.pptx", "ppt_deck_111111");
        var word = registry.RegisterInstance("Word",        "doc.docx",  "word_doc_222222");
        var xl   = registry.RegisterInstance("Excel",       "sheet.xlsx","excel_sheet_333333");

        Assert.Equal(3, new HashSet<string> { ppt, word, xl }.Count);
        Assert.StartsWith("ppt_",   ppt);
        Assert.StartsWith("word_",  word);
        Assert.StartsWith("excel_", xl);
    }
}
