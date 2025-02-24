
When to use SUB, FUNCTION, LET, GET & SET
https://stackoverflow.com/a/75615408

```
> has _return?
  |_ yes > has parameter?
  |        |_ yes: Function
  |        |_ no > likely verb?
  |                |_ yes: Function
  |                |_ no: Get
  |_ no > has parameter?
          |_ yes > likely verb?
          |        |_ yes: Sub
          |        |_ no > uses _object?
          |                |_ yes: Set
          |                |_ no: Let
          |_ no: Sub
```

