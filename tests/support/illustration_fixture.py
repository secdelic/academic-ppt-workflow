"""Synthetic controlled illustration contexts, with no project inputs."""
from scripts.academic_ppt.illustration_requests import PROFILES
SUBJECTS={1:"EDITORIAL_HORIZON",3:"RESEARCH_PROCESS",5:"UNRESOLVED_SEMANTICS",6:"CONCEPTUAL_FRAMEWORK",9:"UNRESOLVED_SEMANTICS",11:"NUMERIC_NATIVE_ANCHOR",12:"IMPLEMENTATION_CUES",14:"CONTINUITY_BACKGROUND",15:"CLOSING_EQUAL_STATEMENTS"}
def context_for(index):
    subject=SUBJECTS[index]
    return {"subject_class":subject,"semantic_boundary":PROFILES[subject]["boundary"],"focal_region":{"x":.65,"y":.1,"w":.3,"h":.8} if PROFILES[subject]["role"]!="NONE" else None,"negative_space_region":{"x":.05,"y":.1,"w":.55,"h":.8},"source_figure_required":False}
