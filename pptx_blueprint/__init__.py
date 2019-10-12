import pathlib
import pptx
from typing import Union, Iterable
from pptx.shapes.base import BaseShape

_Pathlike = Union[str, pathlib.Path]


class Template:
    """Helper class for modifying pptx templates.
    """

    def __init__(self, filename: _Pathlike) -> None:
        """Initializes a Template-Modifier.

        Args:
            filename (path-like): file name or path
        """
        self._template_path = filename
        self._presentation = pptx.Presentation(filename)
        pass

    def replace_text(self, label: str, new_text: str, *, scope=None) -> None:
        """Replaces text placeholders on one or many slides.

        Args:
            label (str): label of the placeholder (without curly braces)
            text (str): new content
            scope: None, slide number, Slide object or iterable of Slide objects
        """
        shapes = self._find_shapes(label)
        
        for shape in shapes:
            shape.text = new_text
        

    def replace_picture(self, label: str, filename: _Pathlike) -> None:
        """Replaces rectangle placeholders on one or many slides.

        Args:
            label (str): label of the placeholder (without curly braces)
            filename (path-like): path to an image file
        """
        pass

    def replace_table(self, label: str, data) -> None:
        """Replaces rectangle placeholders on one or many slides.

        Args:
            label (str): label of the placeholder (without curly braces)
            data (pandas.DataFrame): table to be inserted into the presentation
        """
        pass

    def _find_shapes(self, label: str) -> Iterable[BaseShape]:
        """ Finds all shapes that match the label

        Args:
            label (str): label of the placeholder (without curly braces)
        """
        slide_number, tag_name = label.split(":")
        matched_shapes = []

        def _find_shapes_in_slide(slide):
            for shape in slide.shapes:
                if shape.text == f'{{{tag_name}}}':
                    yield shape

        if slide_number == '*':
            for slide in self._presentation.slides:
                slide_matched_shapes = _find_shapes_in_slide(slide) 
                matched_shapes.extend(slide_matched_shapes)
        else:
            # in label we are using 1 based indexing
            slide_index = int(slide_number)-1
            if slide_index < 0 or slide_index >= len(self._presentation.slides):
                raise IndexError(f"Can't find slide number {slide_number}.")

            slide = self._presentation.slides[slide_index]
            slide_matched_shapes = _find_shapes_in_slide(slide) 
            matched_shapes.extend(slide_matched_shapes)

        return matched_shapes

    def save(self, filename: _Pathlike) -> None:
        """Saves the updated pptx to the specified filepath.

        Args:
            filename (path-like): file name or path
        """
        # TODO: make sure that the user does not override the self._template_path
        pass