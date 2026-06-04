from .location_utils import get_groundup_loss


def get_building_loss(building_value: float, damage_ratio: float):
    """calculate building (Coverage A) loss

    Parameters
    ----------
    building_value : float
        building value (Coverage A)
    damage_ratio : float
        applicable damage ratio
    """
    return get_groundup_loss(building_value, damage_ratio)


def get_appurtenant_loss(appurtenant_value: float, damage_ratio: float):
    """calculate appurtenant (Coverage B) loss

    Parameters
    ----------
    building_value : float
        appurtenant value (Coverage B)
    damage_ratio : float
        applicable damage ratio
    """
    return get_groundup_loss(appurtenant_value, damage_ratio)


def get_content_loss(content_value: float, damage_ratio: float):
    """calculate content (Coverage A) loss

    Parameters
    ----------
    content_value : float
        content value (Coverage C)
    damage_ratio : float
        applicable damage ratio
    """
    return get_groundup_loss(content_value, damage_ratio)


def get_BI_loss(BI_value: float, damage_ratio: float):
    """calculate BI (Coverage D) loss

    Parameters
    ----------
    BI_value : float
        BI value (Coverage D)
    damage_ratio : float
        applicable damage ratio
    """
    return get_groundup_loss(BI_value, damage_ratio)


def get_combined_loss(building_loss: float,
                      appurtenant_loss: float,
                      content_loss: float,
                      BI_loss: float):
    """get combined loss for the location.
    combined loss = A + B + C

    Parameters
    ----------
    building_loss : float
        building loss after damage ratio (Coverage A)
    appurtenant_loss : float
        appurtenant loss after damage ratio (Coverage B)
    content_loss : float
        content loss after damage ratio (Coverage C)
    BI_loss : float
        BI loss after damage ratio (Coverage D)
    """
    return building_loss + appurtenant_loss + content_loss


def get_site_loss(building_loss: float,
                  appurtenant_loss: float,
                  content_loss: float,
                  BI_loss: float):
    """get site loss for the location.
    site loss = A + B + C + D

    Parameters
    ----------
    building_loss : float
        building loss after damage ratio (Coverage A)
    appurtenant_loss : float
        appurtenant loss after damage ratio (Coverage B)
    content_loss : float
        content loss after damage ratio (Coverage C)
    BI_loss : float
        BI loss after damage ratio (Coverage D)
    """
    return building_loss + appurtenant_loss + content_loss + BI_loss
