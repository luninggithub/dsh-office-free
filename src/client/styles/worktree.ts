/** Right-docked Univer panel and sidebar preview styles. */
export const worktreeStyles = `
.uvf_units{display:flex;align-items:center;gap:5px;min-height:36px;padding:5px 9px;border-bottom:1px solid var(--dsw-alias-border-l2,#dfe3eb);overflow-x:auto;scrollbar-width:thin;background:color-mix(in srgb,var(--dsw-alias-bg-base,#fff) 94%,#edf0f8)}
.uvf_unit{display:inline-flex;align-items:center;gap:4px;flex:none;max-width:220px;border:1px solid var(--dsw-alias-border-l2,#dfe3eb);border-radius:8px;background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-secondary,#566174);font:11px/20px inherit;padding:0 9px;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;transition:border-color .12s ease,color .12s ease,background .12s ease}.uvf_unit:hover{border-color:color-mix(in srgb,var(--uvf-accent,#5b6cff) 35%,var(--dsw-alias-border-l2,#dfe3eb));color:var(--dsw-alias-label-primary,#172033)}.uvf_unit_on{border-color:color-mix(in srgb,var(--uvf-accent,#5b6cff) 55%,var(--dsw-alias-border-l2,#dfe3eb));color:#4554d6;background:color-mix(in srgb,var(--uvf-accent,#5b6cff) 9%,var(--dsw-alias-bg-base,#fff));box-shadow:0 0 0 1px color-mix(in srgb,var(--uvf-accent,#5b6cff) 10%,transparent) inset}.uvf_unit_icon{font-size:10px;opacity:.82}
.uvf_viewerShell{position:relative;display:flex;flex:1;min-width:0;min-height:0;background:#f1f3f7}
.uvf_note{display:flex;flex:1;flex-direction:column;align-items:center;justify-content:center;gap:12px;padding:28px;text-align:center;font-size:12px;color:var(--dsw-alias-label-tertiary,#7d8798)}.uvf_note button{border:1px solid color-mix(in srgb,var(--uvf-accent,#5b6cff) 40%,var(--dsw-alias-border-l2,#dfe3eb));border-radius:8px;padding:6px 13px;color:#fff;background:var(--uvf-accent,#5b6cff);font:12px/18px inherit;cursor:pointer}
.uvf_btn{display:grid;place-items:center;width:30px;height:30px;flex:none;padding:0;border:0;border-radius:8px;color:var(--dsw-alias-label-secondary,#566174);background:transparent;cursor:pointer;transition:color .12s ease,background .12s ease,transform .12s ease}.uvf_btn:hover{color:var(--dsw-alias-label-primary,#172033);background:color-mix(in srgb,var(--dsw-alias-label-primary,#172033) 8%,transparent)}.uvf_btn:active{transform:scale(.92)}.uvf_btn:focus-visible{outline:2px solid color-mix(in srgb,var(--uvf-accent,#5b6cff) 65%,transparent);outline-offset:1px}.uvf_btn svg{width:16px;height:16px;fill:none;stroke:currentColor;stroke-width:1.4;stroke-linecap:round;stroke-linejoin:round}
.uvf_dock{--uvf-accent:#5b6cff;position:fixed;z-index:1200;top:calc(8px + var(--dsh-frame-top-clearance,0px));right:0;bottom:0;display:flex;flex-direction:column;min-width:360px;max-width:72vw;overflow:hidden;box-sizing:border-box;color:var(--dsw-alias-label-primary,#172033);background:var(--dsw-alias-bg-base,#fff);border-left:1px solid var(--dsw-alias-border-l2,#dfe3eb);border-radius:14px 0 0 0;box-shadow:-10px 0 32px rgba(13,22,38,.12);pointer-events:auto}
.uvf_dock_dragging{user-select:none}
.uvf_dock_dragging::after{content:"";position:absolute;inset:0;z-index:6;background:transparent}
.uvf_dockSplitter{position:absolute;left:0;top:0;bottom:0;z-index:7;width:10px;margin:0;padding:0;border:0;background:transparent;cursor:col-resize;touch-action:none}
.uvf_dockSplitter::after{content:"";position:absolute;left:0;top:0;bottom:0;width:1px;background:var(--dsw-alias-border-l2,#dfe3eb);transition:background .12s ease,width .12s ease}
.uvf_dockSplitter:hover::after,.uvf_dock_dragging .uvf_dockSplitter::after{width:2px;background:color-mix(in srgb,var(--uvf-accent) 60%,transparent)}
.uvf_dockHead{position:relative;z-index:8;display:flex;align-items:center;gap:8px;height:48px;flex:none;padding:0 8px 0 14px;border-bottom:1px solid var(--dsw-alias-border-l2,#dfe3eb);background:color-mix(in srgb,var(--dsw-alias-bg-base,#fff) 94%,#eef1f8)}
.uvf_dockGlyph{display:grid;place-items:center;width:28px;height:28px;flex:none;border-radius:8px;color:#fff;background:linear-gradient(145deg,#7a87ff,#4f5fe7);box-shadow:0 4px 10px rgba(79,95,231,.22)}
.uvf_dockGlyph svg{width:16px;height:16px;fill:none;stroke:currentColor;stroke-width:1.4;stroke-linecap:round;stroke-linejoin:round}
.uvf_dockIdentity{display:flex;flex:1;min-width:0;flex-direction:column;justify-content:center;line-height:1.2}
.uvf_dockTitle,.uvf_dockFile{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.uvf_dockTitle{font-size:12.5px;font-weight:650}
.uvf_dockFile{margin-top:2px;font-size:10.5px;color:var(--dsw-alias-label-tertiary,#7d8798)}
.uvf_dockChip{display:inline-flex;align-items:center;gap:5px;flex:none;padding:3px 8px;border:1px solid var(--dsw-alias-border-l2,#dfe3eb);border-radius:999px;font-size:10.5px;font-weight:600;line-height:16px;color:#566174}
.uvf_dockDot{width:7px;height:7px;flex:none;border-radius:50%;background:currentColor}
.uvf_dockChip[data-status=draft]{color:#9a6114;border-color:color-mix(in srgb,#d78b25 34%,var(--dsw-alias-border-l2,#dfe3eb));background:color-mix(in srgb,#d78b25 9%,var(--dsw-alias-bg-base,#fff))}
.uvf_dockChip[data-status=ready],.uvf_dockChip[data-status=merged]{color:#16784d;border-color:color-mix(in srgb,#20a66a 30%,var(--dsw-alias-border-l2,#dfe3eb));background:color-mix(in srgb,#20a66a 8%,var(--dsw-alias-bg-base,#fff))}
.uvf_dockChip[data-status=trunk],.uvf_dockChip[data-status=loading],.uvf_dockChip[data-status=discarded]{color:#697386}
.uvf_dockEdit{margin-inline-start:auto}
.uvf_dockChip[data-status=editable]{color:#4554d6;border-color:color-mix(in srgb,var(--uvf-accent,#5b6cff) 40%,var(--dsw-alias-border-l2,#dfe3eb));background:color-mix(in srgb,var(--uvf-accent,#5b6cff) 9%,var(--dsw-alias-bg-base,#fff))}
.uvf_dockChip[data-status=locked]{color:#697386}
.uvf_dockBar{display:flex;align-items:center;gap:6px;flex:none;min-height:36px;padding:4px 10px;overflow-x:auto;scrollbar-width:thin;border-bottom:1px solid var(--dsw-alias-border-l2,#dfe3eb);background:var(--dsw-alias-bg-base,#fff)}
.uvf_dockVersion{flex:none;max-width:170px;height:26px;padding:0 6px;border:1px solid var(--dsw-alias-border-l2,#dfe3eb);border-radius:8px;color:var(--dsw-alias-label-secondary,#566174);background:var(--dsw-alias-bg-base,#fff);font:11px/20px inherit;cursor:pointer}
.uvf_dockAction{flex:none;height:26px;padding:0 10px;border:1px solid var(--dsw-alias-border-l2,#dfe3eb);border-radius:8px;color:var(--dsw-alias-label-secondary,#566174);background:var(--dsw-alias-bg-base,#fff);font:11px/20px inherit;font-weight:600;cursor:pointer;white-space:nowrap;transition:border-color .12s ease,color .12s ease,background .12s ease}
.uvf_dockAction:hover:not(:disabled){border-color:color-mix(in srgb,var(--uvf-accent) 40%,var(--dsw-alias-border-l2,#dfe3eb));color:var(--dsw-alias-label-primary,#172033)}
.uvf_dockAction:disabled{opacity:.55;cursor:default}
.uvf_dockAction:focus-visible{outline:2px solid color-mix(in srgb,var(--uvf-accent) 60%,transparent);outline-offset:1px}
.uvf_dockAction_on{color:#4554d6;border-color:color-mix(in srgb,var(--uvf-accent) 55%,var(--dsw-alias-border-l2,#dfe3eb));background:color-mix(in srgb,var(--uvf-accent) 9%,var(--dsw-alias-bg-base,#fff))}
.uvf_dockAction_primary{color:#fff;border-color:transparent;background:var(--uvf-accent)}
.uvf_dockAction_primary:hover:not(:disabled){color:#fff;background:color-mix(in srgb,var(--uvf-accent) 88%,#000)}
.uvf_dockAction_danger{color:#b73636}.uvf_dockAction_danger:hover:not(:disabled){color:#fff;border-color:transparent;background:#d94242}
.uvf_dockConfirm{display:flex;align-items:center;gap:6px;flex:none;padding:6px 10px;border-bottom:1px solid color-mix(in srgb,#d78b25 30%,var(--dsw-alias-border-l2,#dfe3eb));background:color-mix(in srgb,#d78b25 8%,var(--dsw-alias-bg-base,#fff));font-size:11px;color:#9a6114}
.uvf_dockFailure{flex:none;padding:6px 10px;border-bottom:1px solid color-mix(in srgb,#d94242 28%,var(--dsw-alias-border-l2,#dfe3eb));background:color-mix(in srgb,#d94242 7%,var(--dsw-alias-bg-base,#fff));font-size:11px;color:#b73636}
.uvf_dockBody{display:flex;flex:1;min-height:0;flex-direction:column;background:#f1f3f7}
.uvf_dockBody>.uvf_units{flex:none}
.uvf_dockFrame{display:block;flex:1;min-height:0;width:100%;border:0;background:var(--dsw-alias-bg-base,#fff)}
@media (max-width:680px){.uvf_dock{top:calc(4px + var(--dsh-frame-top-clearance,0px));border-radius:12px 0 0 0}.uvf_dockHead{height:44px;padding-left:12px}}
.uvf_fileView{--uvf-accent:#5b6cff;display:flex;flex:1;flex-direction:column;min-width:0;min-height:0;height:100%;background:var(--dsw-alias-bg-base,#fff)}
.uvf_fileView>.uvf_units{flex:none}
.uvf_fileView .uvf_note{color:var(--dsw-alias-label-tertiary,#7d8798)}
.uvf_surfaceChip{flex:none;align-self:center}
.uvf_fileView>.uvf_dockBody{background:#f1f3f7}
.uvf_noteDetail{max-width:100%;overflow-wrap:anywhere;font:11px/16px ui-monospace,SFMono-Regular,Menlo,monospace;color:var(--dsw-alias-label-tertiary,#7d8798)}
@media (prefers-reduced-motion:reduce){.uvf_dock,.uvf_unit,.uvf_btn,.uvf_dockAction{transition:none}}
`
