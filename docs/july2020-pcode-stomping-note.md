# July 2020 Word Sample: P-code Stomping Note

## Sample

The sample is [`July2020_2485413825.doc.zip`](../samples/July2020_2485413825.doc.zip),
which contains `July2020_2485413825.doc.bin`. The archive was extracted with
the supplied `infected` password. The document was analyzed as data; its macros
were not executed.

Source: https://app.any.run/tasks/ced14735-0ed0-49c4-a4d7-f7fe169a81ee/

## Observed result

The bundled oletools parser found VBA macros and reported VBA stomping. Its only
reported mismatch was the identifier `<crash>`.

pcodedmp emits the instruction `Ld <crash>` in the `ThisDocument` module.
Its `getID()` function maps identifier-table index zero to the first entry in
its `internalNames` list, which is the literal string `<crash>`.

An independent disassembly with pyOpenVBA emitted `Ld name0000` for the same
instruction. pyOpenVBA currently leaves identifier-table operands unresolved,
so this confirms the raw operand is index zero but does not establish what that
entry should mean in this project.

oletools compares names from identifier-bearing P-code instructions against
the VBA source. Because `<crash>` is not a normal VBA identifier and is absent
from the source, this single operand is enough to produce the reported
stomping result.

## Interpretation

This is consistent with unusual, malformed, or stale legacy P-code and could
reflect a limitation in how old P-code is represented or interpreted. It is
not proof that the document was stomped, nor proof that pcodedmp has an
instruction-decoding bug: pyOpenVBA independently reports the same raw
identifier index. Resolving the index against this project's identifier table
and comparing with a trusted Office-generated rebuild would be needed to
determine why the entry is zero.