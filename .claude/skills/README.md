# Skills

Legg prosjektspesifikke skills her. Claude Code oppdager dem automatisk.

Struktur:

```
.claude/skills/
└── min-skill/
    ├── SKILL.md        # påkrevd
    └── ...             # valgfrie hjelpefiler (skript, maler, referanser)
```

Eksempel på `SKILL.md`:

```markdown
---
name: min-skill
description: Hva skillen gjør og når Claude skal bruke den.
---

Instruksjoner til Claude ...
```

Skillen kan så kalles med `/min-skill`, eller Claude bruker den automatisk når beskrivelsen passer.
