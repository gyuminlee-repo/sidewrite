# Journal profiles

A journal profile is one JSON file, `journal_profiles/<name>.json`. It holds what differs
between journals: how a response letter is laid out, how figure and table references are
written, and which colours mark the author's text. `daemon/journal-profile.mjs` is the only
module that reads these files.

## Where a profile is looked up

The first folder that has `<name>.json` wins.

1. `<manuscript folder>/journal_profiles/`
2. each folder in `CC_OFFICE_PROFILE_DIRS` (a PATH-style list, `:` on Linux and macOS)
3. `<CC_OFFICE_CONFIG_DIR>/journal_profiles/` (default `~/.claude/office-addins`)
4. `profiles/journal_profiles/` in this repository

A name is lower case letters, digits, `_` and `-`, at most 64 characters. A file that starts
with `_` is a template and is never listed.

Which profile a paper uses, in order: `journal_profile` in the paper's `revision.json`,
`journal_profile` in a `submission.json` next to the manuscript, the environment variable
`CC_OFFICE_PROFILE`, then `generic`.

## Values of one paper

Do not write measured values of one paper (the colours of this round, its language) into a
profile file. Put them in `profile_overrides`, in `revision.json` or in `submission.json`.
An override of `revision.json` wins over one of `submission.json`.

The merge rule is one level deep: a value that is an object (`colors`, `response_letter`,
`legend_heads`) is merged key by key, every other value (string, list, number) replaces the
profile's value as a whole.

```json
{
  "journal_profile": "example-elsevier",
  "profile_overrides": {
    "colors": { "response_body": "C00000", "response_revised": "0070C0" },
    "response_letter": { "max_words": 250 }
  }
}
```

## Keys

Regular expressions are written for Python `re`. The daemon compiles them as JavaScript and
reports the ones it cannot compile.

| Key                                                           | Meaning                                                                                          |
| ------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `name`                                                        | Profile name, same as the file stem                                                              |
| `reviewer_head`                                               | Regex of a reviewer heading in a letter. Group 1 is the reviewer number                          |
| `comment_head`                                                | Regex of a comment heading in a letter                                                           |
| `ref_head`                                                    | List of regexes, each one a head of a reference (`Fig. `, `Table `, `Supplementary Note `)       |
| `ref_item`                                                    | Regex of one item (number, panel, range)                                                         |
| `ref_sep`                                                     | Regex of the separator between items (`, `, `; `, `and`)                                         |
| `ref_item_cont`                                               | Regex of what may follow a separator in place of an item (one panel letter)                      |
| `ref_caption`                                                 | Regex of a caption title at the start of a paragraph; a reference inside it does not count       |
| `supp_ref_corpus`                                             | Test sentences for the reference grammar: `[sentence, [expected groups]]`                        |
| `body_section_head`                                           | Regex of a manuscript section heading                                                            |
| `colors.response_body`                                        | `RRGGBB` of the author's answer, or null                                                         |
| `colors.response_revised`                                     | `RRGGBB` of location lines and quotes, or null                                                   |
| `colors.response_kr`                                          | `RRGGBB` of the translation paragraph, or null                                                   |
| `colors.highlight`                                            | Word highlight name (`yellow`), or null                                                          |
| `rationale_lang`                                              | Language of the rationale comments. Default: `CC_OFFICE_RATIONALE_LANG`, then `English`          |
| `response_letter.lang`                                        | Language of the response text. Default `English`                                                 |
| `response_letter.reviewer_label`, `editor_label`              | Block headings. Default `Reviewer #{n}`, `Editor`                                                |
| `response_letter.comment_label`                               | Comment heading. Default `Comment {n}`                                                           |
| `response_letter.section_head`                                | Label of a group of comments in the review. Default `{section}`; null writes none                |
| `response_letter.location_line`, `removed_line`               | Default `[Revised in {section}, {lines}]`, `[Removed from {section}, {lines}]`                   |
| `response_letter.declined_line`                               | Default `[No change made]`                                                                       |
| `response_letter.line_one`, `line_range`                      | Default `line {a}`, `lines {a}-{b}`                                                              |
| `response_letter.section_level`                               | Integer: only headings of this level or above name a section. Default null (the nearest heading) |
| `response_letter.rules`                                       | Writing rules given to the model when it drafts a response. Default `[]`                         |
| `response_letter.opening_regex`, `max_words`                  | Checks on each response. Default null                                                            |
| `response_letter.max_quote_words`, `forbid_cross_reviewer`    | Checks on each response. Default null, false                                                     |
| `response_letter.gloss_enabled`, `gloss_lang`, `gloss_prefix` | Translation paragraph after each response. Default false, null, `""`                             |

Keys the daemon does not know are passed through unchanged.

The reference grammar is `ref_head + ref_item + (ref_sep + (ref_item | ref_item_cont))*`.
References joined by a separator alone form one group.

## Profiles in this folder

- `generic.json`: a reference grammar that fits most journals; no colours.
- `example-elsevier.json`: an example for a numbered-section, `Fig. S1`-style journal with a
  test corpus for the reference grammar. Its colours are null: read them from your own
  earlier letter and keep them as overrides of the paper.

## Using an existing profile folder

If you already keep `journal_profiles/<name>.json` files elsewhere, point the daemon at that
folder instead of copying them:

```sh
CC_OFFICE_PROFILE_DIRS=/path/to/journal_profiles npm run dev
```
