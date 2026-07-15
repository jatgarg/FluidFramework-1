# SharedString C# port

## Known limitations

### Obliterate support

`ObliterateRange` supports the non-sided op path: the wire op, local/remote segment marking, visibility filtering, local ack promotion, loading `movedSeq` snapshot stamps, and concurrent-insert eating for inserts whose `refSeq` did not observe an obliterate.

The port can load obliterate snapshot stamps from `movedSeq`/`movedClientIds`, but it does not include a snapshot writer. Sided obliterate and zamboni/summarization parity remain deferred.
