from __future__ import annotations
from typing import List

import uno
from com.sun.star.animations import XAnimateMotion
from com.sun.star.animations import XAnimationNode
from com.sun.star.animations import XAudio
from com.sun.star.animations import XTimeContainer

from ooo.dyn.animations.animation_fill import AnimationFill
from ooo.dyn.beans.named_value import NamedValue
from ooo.dyn.presentation.effect_node_type import EffectNodeTypeEnum
from ooo.dyn.presentation.effect_preset_class import EffectPresetClassEnum

from ooodev.draw import Draw, ImpressDoc
from ooodev.loader import Lo
from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
from ooodev.utils.file_io import FileIO
from ooodev.utils.info import Info


class AnimationDemo:
    def __init__(self) -> None:
        pass

    def show(self) -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:
            doc = ImpressDoc.create_doc(loader)
            try:
                slide = doc.slides[0]  # access first page
                slide.blank_slide()

                # add an ellipse to the center of the slide
                slide_size = slide.get_size_mm()
                width = 50
                height = 50
                x = round((slide_size.Width / 2) - (width / 2))
                y = round((slide_size.Height / 2) - (height / 2))
                s1 = slide.draw_ellipse(x=x, y=y, width=width, height=height)
                try:
                    root = slide.get_animation_node()
                    self._set_user_data(
                        node=root,
                        effect_node_type=EffectNodeTypeEnum.AFTER_PREVIOUS,
                        effect_present_class=EffectPresetClassEnum.MOTIONPATH,
                    )
                    # root --> seq --> par
                    root_time = Lo.qi(XTimeContainer, root, True)
                    seq_time = Lo.create_instance_mcf(
                        XTimeContainer,
                        "com.sun.star.animations.SequenceTimeContainer",
                        raise_err=True,
                    )
                    root_time.appendChild(seq_time)

                    par_time = Lo.create_instance_mcf(
                        XTimeContainer,
                        "com.sun.star.animations.ParallelTimeContainer",
                        raise_err=True,
                    )
                    par_time.Acceleration = 0.05
                    par_time.Decelerate = 0.05
                    par_time.Fill = AnimationFill.HOLD
                    seq_time.appendChild(par_time)

                    # set animation of ellipse to execute in parallel
                    motion = Lo.create_instance_mcf(
                        XAnimateMotion,
                        "com.sun.star.animations.AnimateMotion",
                        raise_err=True,
                    )
                    motion.Duration = 2
                    motion.Fill = AnimationFill.HOLD
                    motion.Target = s1.component
                    motion.Path = "m -0.5 -0.5 0.5 1 0.5 -1"
                    par_time.appendChild(motion)

                    # create audio playing in parallel
                    fnm = Info.get_gallery_dir() / "sounds" / "applause.wav"
                    if fnm.exists() and fnm.is_file():
                        audio = Lo.create_instance_mcf(
                            XAudio, "com.sun.star.animations.Audio", raise_err=True
                        )
                        audio.Source = FileIO.fnm_to_url(fnm)
                        audio.Volume = 1.0
                        par_time.appendChild(audio)
                    else:
                        print(f'Unable to load audio file: "{fnm}"')
                except Exception as e:
                    print(e)

                # slideshow start() crashes if the doc is not visible
                doc.set_visible()

                Lo.delay(500)
                Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
                # show = doc.get_show()
                # show.start() starts slideshow but not necessarily in 100% full screen
                # show.start()
                sc = doc.get_show_controller()
                Draw.wait_ended(sc)
            finally:
                doc.close_doc()

    def _set_user_data(
        self,
        node: XAnimationNode,
        effect_node_type: EffectNodeTypeEnum,
        effect_present_class: EffectPresetClassEnum,
        preset_id: str = "",
    ) -> None:
        if preset_id:
            has_id = True
        else:
            has_id = False

        user_data: List[NamedValue] = [
            NamedValue("node-type", effect_node_type.value),
            NamedValue("preset-class", effect_present_class.value),
        ]
        if has_id:
            user_data.append(NamedValue("preset-id", preset_id))
        node.UserData = tuple(user_data)
        return None
