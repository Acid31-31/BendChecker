namespace BendChecker.Core.Models;

public sealed record StepMeshPart(
    double[] Positions,
    double[] Normals,
    int[] Indices,
    byte Red,
    byte Green,
    byte Blue,
    byte Alpha);
