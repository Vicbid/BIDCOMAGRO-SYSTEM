# BIDCOMAGRO-SYSTEM

Instrucciones del proyecto para Claude Code. El detalle vive en `.ai_agents/` — se importa acá
para que se cargue siempre al arrancar la sesión, sin depender de que alguien lo abra a mano.

@.ai_agents/context.md
@.ai_agents/AI_INSTRUCTIONS.md
@.ai_agents/engineer.md
@.ai_agents/quality.md

## Referencia rápida (por si el import de arriba no resuelve)
- **Nunca** `clasp push`/`deploy` — el usuario deploya manualmente.
- **ES5 estricto** en todo Apps Script (var, function; nada de let/const/=>).
- Bump de `@version` (archivo) + fila en `VERSIONS.md` (raíz) en cada edición funcional.
- Antes de tocar `appsscript.json` (scopes/access/executeAs): avisar que requiere re-auth manual,
  antes de hacer el cambio.
- No renombrar/migrar estados o enums de datos en producción sin preguntar primero.
- Config/administración nueva de cualquier módulo → panel en LAUNCHER, no edición manual de hoja.
- Metodología de seguridad completa (checklist tipo OWASP + patrones de código): skill global
  `project-playbook`.
